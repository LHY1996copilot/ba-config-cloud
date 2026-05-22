from __future__ import annotations

from io import BytesIO
from pathlib import Path
import tempfile

from ba_streamlit_app import (
    choose_generation_upload,
    prepare_downloads,
    requires_manual_process_for_customer_final,
    run_ba_workflow,
    save_uploaded_file,
)


QUOTE_STYLE_OPTIONS = {
    "客户最终报价口径": "customer-final",
    "价格参考表完整报价": "reference",
}


def quote_style_from_label(label: str) -> str:
    return QUOTE_STYLE_OPTIONS[label]


def cloud_resolve_price_path(price_upload, folder: Path, quote_style: str) -> Path | None:
    if quote_style != "reference":
        return None
    if price_upload is None:
        raise FileNotFoundError("选择价格参考表完整报价时，请上传价格参考表。")
    return save_uploaded_file(price_upload, folder, "价格参考表")


def run_cloud_uploaded_generation(
    input_upload,
    price_upload,
    mode: str,
    k: float,
    tag: str,
    input_filename: str,
    quote_style: str,
) -> tuple[list[tuple[str, bytes]], BytesIO]:
    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(tmp)
        input_path = save_uploaded_file(input_upload, work_dir, input_filename)
        price_path = cloud_resolve_price_path(price_upload, work_dir, quote_style)
        output_dir = work_dir / "outputs"
        output_dir.mkdir()
        files = run_ba_workflow(
            input_file=input_path,
            price_file=price_path,
            mode=mode,
            k=k,
            tag=tag,
            output_dir=output_dir,
            quote_style=quote_style,
        )
        return prepare_downloads(files)


def password_gate(st) -> bool:
    password = st.secrets.get("APP_PASSWORD", "")
    if not password:
        return True
    entered = st.text_input("访问密码", type="password")
    if entered == password:
        return True
    if entered:
        st.error("密码不正确。")
    return False


def render_cloud_downloads(prefix: str) -> None:
    import streamlit as st

    downloads = st.session_state.get(f"{prefix}_downloads")
    zip_data = st.session_state.get(f"{prefix}_zip")
    if not downloads or not zip_data:
        return

    st.success("文件已生成。")
    st.download_button(
        "下载全部文件 ZIP",
        data=zip_data,
        file_name="BA系统配置输出.zip",
        mime="application/zip",
        key=f"{prefix}_download_zip",
    )
    for name, data in downloads:
        st.download_button(
            f"下载 {name}",
            data=data,
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{prefix}_{name}",
        )


def render_cloud_app() -> None:
    import streamlit as st

    st.set_page_config(page_title="BA系统配置工具", page_icon="📊", layout="centered")
    st.title("BA系统配置工具")
    st.info("云端版本不会读取本机文件。请按需要上传项目文件；客户最终报价口径不需要价格参考表。")

    if not password_gate(st):
        return

    tab_generate, tab_quote = st.tabs(["生成 DDC / 清单", "生成报价"])

    with tab_generate:
        point_file = st.file_uploader("原始点位表", type=["xlsx"], key="cloud_point_file")
        process_file = st.file_uploader(
            "人工配置过程文档（匹配客户最终报价时必填）",
            type=["xlsx"],
            key="cloud_process_file_generate",
        )
        price_file = st.file_uploader("价格参考表（完整报价口径时必填）", type=["xlsx"], key="cloud_price_file_generate")

        col_k, col_tag = st.columns(2)
        with col_k:
            k_value = st.number_input("K值", min_value=1.0, max_value=2.0, value=1.1, step=0.05, format="%.2f")
        with col_tag:
            tag = st.selectbox("传感器品牌", ["国产", "进口"], key="cloud_tag_generate")

        quote_style_label = st.selectbox("报价口径", list(QUOTE_STYLE_OPTIONS), key="cloud_quote_style_generate")
        quote_style = quote_style_from_label(quote_style_label)

        if st.button("生成文件", type="primary", key="cloud_generate_all"):
            if not point_file:
                st.error("请先上传原始点位表。")
            elif requires_manual_process_for_customer_final(process_file, quote_style):
                st.error("要匹配客户最终报价，请同时上传人工配置过程文档。只上传原始点位表会按自动优化数量计算，结果可能偏低。")
            else:
                workflow_file = choose_generation_upload(point_file, process_file)
                with st.spinner("正在生成，请稍候..."):
                    try:
                        downloads, zip_buffer = run_cloud_uploaded_generation(
                            input_upload=workflow_file,
                            price_upload=price_file,
                            mode="all",
                            k=k_value,
                            tag=tag,
                            input_filename="配置输入文档",
                            quote_style=quote_style,
                        )
                        st.session_state["cloud_generate_downloads"] = downloads
                        st.session_state["cloud_generate_zip"] = zip_buffer.getvalue()
                    except Exception as exc:
                        st.error(f"生成失败：{exc}")

        render_cloud_downloads("cloud_generate")

    with tab_quote:
        list_file = st.file_uploader("人工处理后的清单文档", type=["xlsx"], key="cloud_list_file")
        price_file = st.file_uploader("价格参考表（完整报价口径时必填）", type=["xlsx"], key="cloud_price_file_quote")
        tag = st.selectbox("传感器品牌", ["国产", "进口"], key="cloud_tag_quote")
        quote_style_label = st.selectbox("报价口径", list(QUOTE_STYLE_OPTIONS), key="cloud_quote_style_quote")
        quote_style = quote_style_from_label(quote_style_label)

        if st.button("生成报价", type="primary", key="cloud_generate_quote"):
            if not list_file:
                st.error("请先上传人工处理后的清单文档。")
            else:
                with st.spinner("正在生成，请稍候..."):
                    try:
                        downloads, zip_buffer = run_cloud_uploaded_generation(
                            input_upload=list_file,
                            price_upload=price_file,
                            mode="quote",
                            k=1.1,
                            tag=tag,
                            input_filename="人工处理后清单",
                            quote_style=quote_style,
                        )
                        st.session_state["cloud_quote_downloads"] = downloads
                        st.session_state["cloud_quote_zip"] = zip_buffer.getvalue()
                    except Exception as exc:
                        st.error(f"生成失败：{exc}")

        render_cloud_downloads("cloud_quote")


if __name__ == "__main__":
    render_cloud_app()
