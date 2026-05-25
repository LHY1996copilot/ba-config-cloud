# BA系统配置工具云端版

Streamlit Cloud 入口文件：`ba_cloud_app.py`

## 使用方式

- 上传原始点位表。
- 设置 K 值和传感器品牌。
- 点击“生成清单和报价”，下载清单文档和报价文档。

## 部署

在 Streamlit Community Cloud 创建应用时：

- Repository：当前仓库
- Branch：`main`
- Main file path：`ba_cloud_app.py`

可选访问密码：在 Streamlit Cloud Secrets 中设置：

```toml
APP_PASSWORD = "你的访问密码"
```
