from __future__ import annotations

from io import BytesIO
from pathlib import Path
import tempfile

from ba_streamlit_app import (
    prepare_downloads,
    run_ba_workflow,
    save_uploaded_file,
)


SIMPLE_OUTPUT_NAMES = {"清单文档.xlsx", "报价文档.xlsx"}


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


def run_simple_uploaded_generation(input_upload, k: float, tag: str) -> tuple[list[tuple[str, bytes]], BytesIO]:
    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(tmp)
        input_path = save_uploaded_file(input_upload, work_dir, "原始点位表")
        output_dir = work_dir / "outputs"
        output_dir.mkdir()
        files = run_ba_workflow(
            input_file=input_path,
            price_file=None,
            mode="all",
            k=k,
            tag=tag,
            output_dir=output_dir,
            quote_style="customer-final",
        )
        customer_files = [path for path in files if path.name in SIMPLE_OUTPUT_NAMES]
        return prepare_downloads(customer_files)


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
    st.info("上传原始点位表，设置 K 值和传感器品牌，即可生成清单文档和报价文档。")

    if not password_gate(st):
        return

    point_file = st.file_uploader("原始点位表", type=["xlsx"], key="cloud_point_file")
    col_k, col_tag = st.columns(2)
    with col_k:
        k_value = st.number_input("K值", min_value=1.0, max_value=2.0, value=1.1, step=0.05, format="%.2f")
    with col_tag:
        tag = st.selectbox("传感器品牌", ["国产", "进口"], key="cloud_tag_generate")

    if st.button("生成清单和报价", type="primary", key="cloud_generate_all"):
        if not point_file:
            st.error("请先上传原始点位表。")
        else:
            with st.spinner("正在生成，请稍候..."):
                try:
                    downloads, zip_buffer = run_simple_uploaded_generation(point_file, k_value, tag)
                    st.session_state["cloud_generate_downloads"] = downloads
                    st.session_state["cloud_generate_zip"] = zip_buffer.getvalue()
                except Exception as exc:
                    st.error(f"生成失败：{exc}")

    render_cloud_downloads("cloud_generate")


if __name__ == "__main__":
    render_cloud_app()
