# 水电表计费系统

一个基于 Rust 开发的智能水电表计费系统，支持多电表、自动计算、专业 Word 文档生成，以及 Web 界面上传功能。

## ✨ 核心功能

### 🔌 多电表支持
- 支持同一商户配置多个电表
- 自动识别电表列（电表1上期读数、电表1本期读数等）
- 智能计算总用电量和电费

### 💰 智能计费计算
- **水费计算**：四舍五入到元
- **电费计算**：先合计总用电量，再乘单价，最后四舍五入到元
- **其他费用**：水电人工费、垃圾处理费、滞纳金、广告费
- **总费用**：各费用项相加得出最终金额

### 📄 专业 Word 文档生成
- 自定义标题（支持自定义或默认格式）
- 表格化布局，数据清晰易读
- 所有数字居中对齐
- 中文大写金额显示
- 支持分页设置（每页显示指定数量的商户表格）

### 🌐 Web 服务界面
- 简洁的上传页面
- 支持 Excel (.xlsx) 和 CSV 文件
- 实时生成并下载 Word 文档
- 可配置抄表人、抄表日期等参数

## 📊 数据格式要求

### 必需表头字段
CSV/Excel 文件必须包含以下表头（顺序不强制）：

| 字段名 | 说明 | 示例 |
|--------|------|------|
| 铺面编号 | 商户编号 | PM-001 |
| 店铺名称 | 商户名称 | 张三商店 |
| 上期水表读数 | 上月水表读数 | 500 |
| 本期水表读数 | 本月水表读数 | 508 |
| 水费单价 | 水费单价（元/吨） | 1.1180 |
| 电费单价 | 电费单价（元/度） | 1.0300 |
| 水电人工费 | 人工服务费 | 50.00 |
| 垃圾处理费 | 垃圾处理费用 | 20.00 |
| 电表1上期读数 | 电表1上月读数 | 5063 |
| 电表1本期读数 | 电表1本月读数 | 5809 |
| 电表2上期读数 | 电表2上月读数 | 1200 |
| 电表2本期读数 | 电表2本月读数 | 1280 |

### 示例数据行
```csv
铺面编号,店铺名称,上期水表读数,本期水表读数,水费单价,电费单价,电表1上期读数,电表1本期读数,电表2上期读数,电表2本期读数,水电人工费,垃圾处理费
PM-001,张三商店,500,508,1.1180,1.0300,5063,5809,1200,1280,50,20
```

## 🚀 快速开始

### 环境要求
- Rust 1.70+
- macOS/Linux/Windows

### 本地运行

1. **克隆项目**
```bash
git clone <repository-url>
cd water_and_electricity_meter
```

2. **编译项目**
```bash
cargo build --release
```

3. **启动 Web 服务**
```bash
# 默认端口 3002
./target/release/server

# 或指定端口
PORT=3001 ./target/release/server
```

4. **访问服务**
- 浏览器打开：`http://localhost:3002/`
- 上传 CSV/Excel 文件生成 Word 文档

### 后台运行
```bash
# 使用 nohup 后台运行
nohup ./target/release/server > server.log 2>&1 &

# 查看日志
tail -f server.log

# 停止服务
pkill -f "target/release/server"
```

### 使用部署脚本
```bash
# 使用项目自带的部署脚本
./deploy.sh
```

## 🐳 Docker 部署

### 使用 Docker Compose（推荐）
```bash
# 构建并启动
docker compose up -d --build

# 查看日志
docker compose logs -f

# 停止服务
docker compose down
```

### 手动 Docker 构建
```bash
# 构建镜像
docker build -t water-meter .

# 运行容器
docker run -d -p 3002:3002 --name water-meter water-meter
```

## 📁 项目结构

```
water_and_electricity_meter/
├── src/
│   ├── lib.rs              # 核心库：数据结构、Word生成、文件解析
│   ├── server.rs           # Web 服务：上传页面、文件处理
│   ├── main.rs             # CLI 工具：命令行生成 Word
│   └── generate_sample.rs  # 示例数据生成器
├── Cargo.toml              # 项目配置和依赖
├── Dockerfile              # Docker 构建文件
├── docker-compose.yml      # Docker Compose 配置
├── deploy.sh               # 本地部署脚本
└── README.md               # 项目说明文档
```

## 🔧 配置选项

### Web 界面配置
- **自定义标题**：可设置文档标题，默认格式为"yyyy年MM月抄表计费通知单"
- **每页表格数量**：控制 Word 文档中每页显示的商户表格数量
- **抄表人**：设置抄表人员姓名
- **抄表日期**：设置抄表日期

### 生成选项
```rust
pub struct GenerateOptions {
    pub custom_title: Option<String>,  // 自定义标题
    pub per_page: usize,               // 每页表格数量
}
```

## 📋 生成的 Word 文档格式

### 文档结构
1. **标题**：自定义或默认标题（24号字体，居中，加粗）
2. **基本信息**：编号、姓名、抄表人、抄表日期
3. **费用明细表格**：
   - 表头：项目、上月表底、本月抄表数、实用度数、公共分摊、单价（元）、金额
   - 电表行：逐表显示用电量和费用
   - 水费行：水表读数和费用
   - 其他费用：水电人工费、垃圾处理费、滞纳金、广告费
   - 合计行：总费用（中文大写 + 小写金额）
4. **说明文字**：收费规则和注意事项

### 表格特点
- 所有数字居中对齐
- 多电表时单价和金额列自动合并
- 金额列宽度优化
- 支持分页和分隔符

## 🛠️ 开发说明

### 主要数据结构
```rust
pub struct MerchantBill {
    pub merchant_name: String,           // 商户名称
    pub shop_code: String,              // 铺面编号
    pub electricity_meters: Vec<ElectricityMeter>, // 电表列表
    pub water_usage: f64,               // 用水量
    pub electricity_usage: f64,         // 总用电量
    pub total_fee: f64,                 // 总费用
    // ... 其他字段
}

pub struct ElectricityMeter {
    pub meter_id: String,               // 电表ID
    pub prev_reading: f64,              // 上期读数
    pub curr_reading: f64,              // 本期读数
    pub usage: f64,                     // 用电量
    pub amount: f64,                    // 电费金额
}
```

### 核心函数
- `read_data_file()`: 解析 Excel/CSV 文件
- `generate_word_document_with_template()`: 生成 Word 文档
- `find_electricity_columns()`: 动态识别电表列
- `rmb_upper()`: 金额转中文大写

## 🚨 注意事项

1. **文件格式**：支持 .xlsx 和 .csv 格式，第一行必须是表头
2. **数据完整性**：确保电表列成对出现（上期读数 + 本期读数）
3. **金额精度**：水费四舍五入到元，电费四舍五入到元
4. **端口配置**：默认端口 3002，可通过环境变量 PORT 修改

## 📝 更新日志

### v0.1.0
- ✅ 基础多电表支持
- ✅ Excel/CSV 文件解析
- ✅ 专业 Word 文档生成
- ✅ Web 界面上传功能
- ✅ Docker 部署支持
- ✅ 智能计费计算
- ✅ 表格居中对齐
- ✅ 中文大写金额
- ✅ 滞纳金、广告费占位行

## 📄 许可证

本项目仅作为示例/内部工具使用，未附加开源协议。根据实际需要补充许可声明。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request 来改进项目功能。

---

**快速测试**：访问 `http://localhost:3002/` 上传示例 CSV 文件，体验完整的 Word 文档生成流程。
