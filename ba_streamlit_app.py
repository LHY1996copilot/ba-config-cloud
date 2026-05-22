from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
import tempfile
from uuid import uuid4
import zipfile

from ba_config import run as run_cli_workflow


OUTPUT_ORDER = ["DDC配置文档.xlsx", "清单文档.xlsx", "报价文档.xlsx"]
DEFAULT_PRICE_FILENAME = "价格参考表0520.xlsx"
WEB_OUTPUT_ROOT = Path(__file__).resolve().parent / "ba_config_outputs" / "web"


def find_default_price_file() -> Path | None:
    search_roots = [Path.cwd(), Path(__file__).resolve().parent]
    for root in search_roots:
        candidate = root / DEFAULT_PRICE_FILENAME
        if candidate.exists():
            return candidate

    users_root = Path("E:/Users")
    if users_root.exists():
        matches = sorted(users_root.glob(f"*/Desktop/0520*/{DEFAULT_PRICE_FILENAME}"))
        for match in matches:
            if match.is_file() and not match.name.startswith("~$"):
                return match
    return None


def collect_excel_outputs(output_dir: Path) -> list[Path]:
    files = [path for path in output_dir.glob("*.xlsx") if path.is_file()]
    order = {name: index for index, name in enumerate(OUTPUT_ORDER)}
    return sorted(files, key=lambda path: (order.get(path.name, len(order)), path.name))


def create_web_output_dir(root: Path = WEB_OUTPUT_ROOT, mode: str = "run") -> Path:
    root.mkdir(parents=True, exist_ok=True)
    safe_mode = "".join(char for char in mode if char.isalnum() or char in {"-", "_"}) or "run"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = root / f"{stamp}_{safe_mode}_{uuid4().hex[:6]}"
    output_dir.mkdir()
    return output_dir


def make_zip_bytes(files: list[Path]) -> BytesIO:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, arcname=path.name)
    buffer.seek(0)
    return buffer


def run_ba_workflow(
    input_file: Path,
    price_file: Path | None,
    mode: str,
    k: float,
    tag: str,
    output_dir: Path,
    ddc_source: str = "auto",
    quote_style: str = "reference",
) -> list[Path]:
    class Args:
        pass

    args = Args()
    args.input = str(input_file)
    args.price = str(price_file) if price_file else ""
    args.mode = mode
    args.k = k
    args.tag = tag
    args.output_dir = str(output_dir)
    args.ddc_source = ddc_source
    args.quote_style = quote_style
    run_cli_workflow(args)
    return collect_excel_outputs(output_dir)


def file_bytes(path: Path) -> bytes:
    return path.read_bytes()


def save_uploaded_file(uploaded_file, folder: Path, filename: str) -> Path:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    path = folder / f"{filename}{suffix}"
    path.write_bytes(uploaded_file.getbuffer())
    return path


def resolve_price_path(price_upload, folder: Path) -> Path:
    if price_upload is not None:
        return save_uploaded_file(price_upload, folder, "价格参考表")
    default_price = find_default_price_file()
    if default_price is None:
        raise FileNotFoundError("未找到价格参考表，请上传价格参考表0520.xlsx。")
    return default_price


def prepare_downloads(files: list[Path]) -> tuple[list[tuple[str, bytes]], BytesIO]:
    items = [(path.name, file_bytes(path)) for path in files]
    return items, make_zip_bytes(files)


def choose_generation_upload(point_upload, process_upload):
    return process_upload if process_upload is not None else point_upload


def render_app() -> None:
    import streamlit as st

    st.set_page_config(page_title="BA系统配置工具", page_icon="📊", layout="centered")
    st.title("BA系统配置工具")
    default_price = find_default_price_file()
    if default_price:
        st.info(f"已检测到价格参考表：{default_price.name}。不上传价格参考表时，会自动使用这份文件。")
    else:
        st.warning("未检测到默认价格参考表。生成文件前请上传价格参考表。")

    tab_generate, tab_quote = st.tabs(["生成 DDC / 清单", "生成报价"])

    with tab_generate:
        point_file = st.file_uploader("原始点位表", type=["xlsx"], key="point_file")
        process_file = st.file_uploader(
            "人工配置过程文档（可选，推荐用于匹配最终报价）",
            type=["xlsx"],
            key="process_file_generate",
        )
        price_file = st.file_uploader("价格参考表（可选）", type=["xlsx"], key="price_file_generate")
        col_k, col_tag = st.columns(2)
        with col_k:
            k_value = st.number_input("K值", min_value=1.0, max_value=2.0, value=1.1, step=0.05, format="%.2f")
        with col_tag:
            tag = st.selectbox("传感器品牌", ["国产", "进口"], key="tag_generate")
        quote_style_label = st.selectbox(
            "报价口径",
            ["客户最终报价口径", "价格参考表完整报价"],
            key="quote_style_generate",
        )
        quote_style = "customer-final" if quote_style_label == "客户最终报价口径" else "reference"

        if st.button("生成文件", type="primary", key="generate_all"):
            if not point_file:
                st.error("请先上传原始点位表。")
            else:
                workflow_file = choose_generation_upload(point_file, process_file)
                if process_file is None and quote_style == "customer-final":
                    st.warning("未上传人工配置过程文档，本次会按原始点位自动优化模块数量，报价可能与人工最终报价不同。")
                with st.spinner("正在生成，请稍候..."):
                    try:
                        downloads, zip_buffer, saved_dir = _run_uploaded_generation(
                            workflow_file,
                            price_file,
                            mode="all",
                            k=k_value,
                            tag=tag,
                            point_filename="配置输入文档",
                            price_filename="价格参考表",
                            quote_style=quote_style,
                        )
                        st.session_state["generate_downloads"] = downloads
                        st.session_state["generate_zip"] = zip_buffer.getvalue()
                        st.session_state["generate_saved_dir"] = str(saved_dir)
                    except Exception as exc:
                        st.error(f"生成失败：{exc}")

        _render_downloads("generate")

    with tab_quote:
        list_file = st.file_uploader("人工处理后的清单文档", type=["xlsx"], key="list_file")
        price_file = st.file_uploader("价格参考表（可选）", type=["xlsx"], key="price_file_quote")
        tag = st.selectbox("传感器品牌", ["国产", "进口"], key="tag_quote")
        quote_style_label = st.selectbox(
            "报价口径",
            ["客户最终报价口径", "价格参考表完整报价"],
            key="quote_style_quote",
        )
        quote_style = "customer-final" if quote_style_label == "客户最终报价口径" else "reference"

        if st.button("生成报价", type="primary", key="generate_quote"):
            if not list_file:
                st.error("请先上传人工处理后的清单文档。")
            else:
                with st.spinner("正在生成，请稍候..."):
                    try:
                        downloads, zip_buffer, saved_dir = _run_uploaded_generation(
                            list_file,
                            price_file,
                            mode="quote",
                            k=1.1,
                            tag=tag,
                            point_filename="人工处理后清单",
                            price_filename="价格参考表",
                            quote_style=quote_style,
                        )
                        st.session_state["quote_downloads"] = downloads
                        st.session_state["quote_zip"] = zip_buffer.getvalue()
                        st.session_state["quote_saved_dir"] = str(saved_dir)
                    except Exception as exc:
                        st.error(f"生成失败：{exc}")

        _render_downloads("quote")


def _run_uploaded_generation(
    input_upload,
    price_upload,
    mode: str,
    k: float,
    tag: str,
    point_filename: str,
    price_filename: str,
    quote_style: str,
) -> tuple[list[tuple[str, bytes]], BytesIO, Path]:
    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(tmp)
        input_path = save_uploaded_file(input_upload, work_dir, point_filename)
        price_path = resolve_price_path(price_upload, work_dir) if quote_style == "reference" else None
        output_dir = create_web_output_dir(mode=mode)
        files = run_ba_workflow(
            input_file=input_path,
            price_file=price_path,
            mode=mode,
            k=k,
            tag=tag,
            output_dir=output_dir,
            quote_style=quote_style,
        )
        downloads, zip_buffer = prepare_downloads(files)
        return downloads, zip_buffer, output_dir


def _render_downloads(prefix: str) -> None:
    import streamlit as st

    downloads = st.session_state.get(f"{prefix}_downloads")
    zip_data = st.session_state.get(f"{prefix}_zip")
    saved_dir = st.session_state.get(f"{prefix}_saved_dir")
    if not downloads or not zip_data:
        return

    st.success("文件已生成。")
    if saved_dir:
        st.info("文件也已保存到本地文件夹。下载按钮没有反应时，请在电脑中打开这个文件夹取文件。")
        st.code(saved_dir, language="text")
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


if __name__ == "__main__":
    render_app()
