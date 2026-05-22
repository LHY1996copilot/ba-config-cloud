# BA系统配置工具云端版

Streamlit Cloud 入口文件：`ba_cloud_app.py`

## 使用方式

- 客户最终报价口径：不需要上传价格参考表，但必须上传人工配置过程文档，否则会按自动优化数量计算，不能匹配最终报价。
- 价格参考表完整报价：必须上传价格参考表。
- 如需匹配已经发给客户的最终报价，请同时上传原始点位表和人工配置过程文档。

## 部署

在 Streamlit Community Cloud 创建应用时：

- Repository：当前仓库
- Branch：`main`
- Main file path：`ba_cloud_app.py`

可选访问密码：在 Streamlit Cloud Secrets 中设置：

```toml
APP_PASSWORD = "你的访问密码"
```
