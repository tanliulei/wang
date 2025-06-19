# PDF转换神器 - 云部署版

一个强大的PDF到Excel转换工具，支持微信支付明细等PDF文件的智能转换和数据处理。

## 功能特性

- 📄 PDF文件智能解析
- 📊 自动转换为Excel格式
- 🔄 数据自动处理和排序
- 📱 响应式Web界面
- ☁️ 云端部署，随时访问

## 在线使用

访问部署的应用地址即可开始使用。

## 本地运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行应用
streamlit run app.py
```

## 云部署选项

### 1. Streamlit Cloud (推荐)
- 免费且简单
- 直接连接GitHub仓库
- 自动部署

### 2. Railway
- 支持Dockerfile
- 免费额度充足
- 部署简单

### 3. Heroku
- 经典PaaS平台
- 使用Procfile部署
- 免费层有限制

### 4. Docker部署
```bash
# 构建镜像
docker build -t pdf-converter .

# 运行容器
docker run -p 8501:8501 pdf-converter
```

## 技术栈

- **前端**: Streamlit
- **后端**: Python
- **PDF处理**: pdfplumber
- **Excel处理**: openpyxl, pandas
- **容器化**: Docker

## 使用说明

1. 上传PDF文件
2. 自动解析和处理
3. 下载生成的Excel文件

## 注意事项

- 支持最大200MB的PDF文件
- 确保PDF包含结构化数据
- 处理大文件可能需要较长时间
