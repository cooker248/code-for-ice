import pandas as pd
import os
import streamlit as st
import io
import zipfile


@st.cache_data
def read_confirm_data(confirm_file):
    confirm_data = pd.read_excel(confirm_file, dtype={"批单序号": str, "险种代码": str})
    return confirm_data


@st.cache_data
def read_baodan_data(baodan_file):
    baodan_data = pd.read_excel(baodan_file, dtype={"批单号": str, "险种代码": str})

    return baodan_data


@st.cache_data
def match_people(baodan_data, confirm_data):
    """
    匹配投保人和被投保人
    """
    # 提取投保人信息
    people_data = baodan_data[["保单号", "投保人", "被保人名称"]]
    people_data = people_data.drop_duplicates(subset=["保单号"])
    # 匹配投保人信息
    merged_data = pd.merge(
        left=confirm_data, right=people_data, on="保单号", how="left"
    )

    # 只提取有用的列
    useful_columns = [
        "保单号",
        "批单序号",
        "投保人",
        "被保人名称",
        "缴费期次",
        "险种代码",
        "险种名称",
        "归属机构",
        "渠道",
        "承保确认时间",
        "费用计提时间",
        "核保时间",
        "保险起期",
        "实收时间",
        "业务员",
        "总费用比例(%)",
        "总费用金额",
        "手续费比例(%)",
        "手续费金额",
        "展业费比例(%)",
        "展业费金额",
        "绩效提奖比例(%)",
        "绩效提奖金额",
        "保费",
        "干预状态",
    ]

    extracted_data = merged_data.loc[:, useful_columns]

    return extracted_data


@st.cache_data
def match_insurance_fees(extracted_data, baodan_data):
    """
    Match the insurance fee and put it in the last columns
    """
    # 修改列标签
    extracted_data.rename(columns={"干预状态": "保额（仅诉责）"}, inplace=True)

    fee_data = baodan_data[["保单号", "批单号", "险种代码", "保险金额"]].rename(
        columns={"批单号": "批单序号"}
    )
    fee_data.drop_duplicates(subset=["保单号", "批单序号", "险种代码"], inplace=True)
    # 合并两个数据表
    suze_df = pd.merge(
        left=extracted_data,
        right=fee_data,
        on=["保单号", "批单序号", "险种代码"],
        how="left",
        validate="many_to_one",
        indicator=True,
    )
    # 确保匹配正确
    # if extracted_data.shape[0] != suze_df.shape[0]:
    #     st.write("匹配有问题")
    unmatched_rows = suze_df[
        (suze_df["_merge"] == "left_only") & (suze_df["险种代码"] == "0460")
    ]
    if not unmatched_rows.empty:
        st.write("以下保单号在保费清单表未找到匹配项：")
        st.write(unmatched_rows["保单号"])
    suze_df.drop(columns="_merge", inplace=True)

    suze_df["保额（仅诉责）"] = suze_df.apply(
        lambda row: (
            row["保险金额"] if row["险种代码"] == "0460" else row["保额（仅诉责）"]
        ),
        axis=1,
    )

    # 处理为0的列
    mask = (suze_df["险种代码"] == "0460") & (suze_df["保额（仅诉责）"] == 0)
    replace_values = fee_data.loc[(fee_data["批单序号"] == "000") & mask, "保险金额"]
    suze_df.loc[mask, "保额（仅诉责）"] = replace_values.astype(float).values

    suze_df.drop(columns="保险金额", inplace=True)

    return suze_df


def create_excel_download_button(df, label, file_name, index=False):
    excel_data = io.BytesIO()
    with pd.ExcelWriter(excel_data, engine="openpyxl") as writer:
        df.to_excel(writer, index=index)
    excel_data.seek(0)
    st.download_button(
        label=label,
        data=excel_data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@st.cache_data
def create_zip(suze_df, pivot_df):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for company_name, group_df in pivot_df.groupby(level=0):
            # Create a new Excel file for each group
            file_buffer = io.BytesIO()

            # Write data to the first sheet
            data_df = suze_df[suze_df["归属机构"] == company_name]
            with pd.ExcelWriter(file_buffer, engine="openpyxl") as writer:
                data_df.to_excel(writer, sheet_name="Sheet1", index=False)

            # Write data to the second sheet
            with pd.ExcelWriter(file_buffer, engine="openpyxl", mode="a") as writer:
                group_df.to_excel(writer, sheet_name="Sheet2", index=True)

            # Add the file to the zip archive
            file_buffer.seek(0)
            zip_file.writestr(f"part/{company_name}.xlsx", file_buffer.read())

            # Close and clear file buffer to release memory
            file_buffer.close()

    # Seek to the beginning of the zip buffer
    zip_buffer.seek(0)
    return zip_buffer


def main():
    st.title("fee process")

    # 允许用户上传Excel文件
    confirm_file = st.file_uploader("上传费用查询表", type=["xlsx", "xls"])
    baodan_file = st.file_uploader("上传保费清单表", type=["xlsx", "xls"])
    if confirm_file is not None and baodan_file is not None:
        try:
            confirm_data = read_confirm_data(confirm_file)
            baodan_data = read_baodan_data(baodan_file)

            extracted_data = match_people(baodan_data, confirm_data)
            suze_df = match_insurance_fees(extracted_data, baodan_data)

            pivot_df = pd.pivot_table(
                suze_df,
                index=["归属机构", "业务员"],
                values=["展业费金额", "绩效提奖金额"],
                aggfunc="sum",
            )

            create_excel_download_button(
                suze_df, "下载费用查询确认表", "费用查询确认.xlsx"
            )
            create_excel_download_button(pivot_df, "下载数据透视表", "数据透视表.xlsx", index=True)

            zip_buffer = create_zip(suze_df, pivot_df)
            # Download button
            button_label = "Download Zip File"
            if st.download_button(
                button_label, data=zip_buffer, file_name="ExcelFiles.zip"
            ):
                st.success(f"{button_label} successfully initiated!")

        except Exception as e:
            st.error("发生错误: {}".format(e))
            st.error("出错了，请确保上传文件正确")
    else:
        st.write("确保上传文件正确")


if __name__ == "__main__":
    main()
