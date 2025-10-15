use anyhow::{Context, Result};
use clap::{Parser, Subcommand};
use std::path::Path;
use calamine::{open_workbook, DataType, Reader, Xlsx};
use chrono::{Datelike, Local};
use std::fs::File;
use std::io::{BufRead, BufReader};

// 导入模板模块
mod template_simple;

#[derive(Debug, Clone)]
pub struct ElectricityMeter {
    pub meter_id: String,
    pub prev_reading: f64,
    pub curr_reading: f64,
    pub usage: f64,
    pub amount: f64,
}

#[derive(Debug, Clone)]
pub struct MerchantBill {
    pub merchant_name: String,
    pub electricity_meters: Vec<ElectricityMeter>,
    pub prev_water_reading: f64,
    pub curr_water_reading: f64,
    pub water_unit_price: f64,
    pub electricity_unit_price: f64,
    pub water_usage: f64,
    pub water_amount: f64,
    pub electricity_usage: f64,  // 总用电量
    pub electricity_amount: f64, // 总电费
    pub water_electricity_labor_fee: f64, // 水电人工费
    pub garbage_disposal_fee: f64, // 垃圾处理费
    pub total_fee: f64,
    pub month: String,
}

impl MerchantBill {
    pub fn new(merchant_name: String, water_unit_price: f64, electricity_unit_price: f64) -> Self {
        Self {
            merchant_name,
            electricity_meters: Vec::new(),
            prev_water_reading: 0.0,
            curr_water_reading: 0.0,
            water_unit_price,
            electricity_unit_price,
            water_usage: 0.0,
            water_amount: 0.0,
            electricity_usage: 0.0,
            electricity_amount: 0.0,
            water_electricity_labor_fee: 0.0,
            garbage_disposal_fee: 0.0,
            total_fee: 0.0,
            month: Local::now().format("%Y年%m月").to_string(),
        }
    }

    pub fn add_electricity_meter(&mut self, meter_id: String, prev_reading: f64, curr_reading: f64) {
        let usage = (curr_reading - prev_reading).max(0.0);
        let amount = usage * self.electricity_unit_price;
        
        let meter = ElectricityMeter {
            meter_id,
            prev_reading,
            curr_reading,
            usage,
            amount,
        };
        
        self.electricity_meters.push(meter);
        self.update_totals();
    }

    pub fn set_water_readings(&mut self, prev_reading: f64, curr_reading: f64) {
        self.prev_water_reading = prev_reading;
        self.curr_water_reading = curr_reading;
        self.water_usage = (curr_reading - prev_reading).max(0.0);
        self.water_amount = self.water_usage * self.water_unit_price;
        self.update_totals();
    }

    pub fn set_additional_fees(&mut self, water_electricity_labor_fee: f64, garbage_disposal_fee: f64) {
        self.water_electricity_labor_fee = water_electricity_labor_fee;
        self.garbage_disposal_fee = garbage_disposal_fee;
        self.update_totals();
    }

    fn update_totals(&mut self) {
        self.electricity_usage = self.electricity_meters.iter().map(|m| m.usage).sum();
        self.electricity_amount = self.electricity_meters.iter().map(|m| m.amount).sum();
        self.total_fee = self.electricity_amount + self.water_amount + self.water_electricity_labor_fee + self.garbage_disposal_fee;
    }

    // 获取所有电表的详细信息（用于模板显示）
    pub fn get_electricity_details(&self) -> Vec<String> {
        self.electricity_meters.iter().map(|meter| {
            format!("电表{}: 上期{}度, 本期{}度, 用量{}度, 费用{:.2}元", 
                meter.meter_id, 
                meter.prev_reading, 
                meter.curr_reading, 
                meter.usage, 
                meter.amount)
        }).collect()
    }
}

#[derive(Clone)]
pub struct HeadersMap<'a> {
    pub merchant: &'a str,
    pub water_prev: &'a str,
    pub water_curr: &'a str,
    pub water_price: &'a str,
    pub electricity_price: &'a str,
    // 电表列的前缀，支持多个电表
    pub electricity_prefix: &'a str,
}

impl<'a> HeadersMap<'a> {
    fn to_pairs(&self) -> Vec<(&str, &str)> {
        vec![
            (self.merchant, "店铺名称"),
            (self.water_prev, "上期水表读数"),
            (self.water_curr, "本期水表读数"),
            (self.water_price, "水费单价"),
            (self.electricity_price, "电费单价"),
        ]
    }
}

#[derive(Parser)]
#[command(name = "excel_to_word")]
#[command(about = "将Excel/CSV数据转换为Word文档")]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// 使用配置文件生成Word文档
    Config {
        /// 输入文件路径
        #[arg(short, long)]
        input: String,
        /// 输出文件路径
        #[arg(short, long)]
        output: String,
        /// 配置文件路径
        #[arg(short, long)]
        config: String,
    },
    /// 使用默认配置生成Word文档
    Default {
        /// 输入文件路径
        #[arg(short, long)]
        input: String,
        /// 输出文件路径
        #[arg(short, long)]
        output: String,
    },
    /// 使用传统方式生成Word文档
    Legacy {
        /// 输入文件路径
        #[arg(short, long)]
        input: String,
        /// 输出文件路径
        #[arg(short, long)]
        output: String,
    },
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    match &cli.command {
        Commands::Config { input, output, config } => {
            println!("使用配置文件生成Word文档...");
            let bills = read_data_file(input, &get_default_headers())?;
            let docx_content = generate_word_document_with_template(&bills, Some(config))?;
            std::fs::write(output, docx_content)?;
            println!("✅ Word文档生成成功: {}", output);
        }
        Commands::Default { input, output } => {
            println!("使用默认配置生成Word文档...");
            let bills = read_data_file(input, &get_default_headers())?;
            let docx_content = generate_word_document_with_template(&bills, None)?;
            std::fs::write(output, docx_content)?;
            println!("✅ Word文档生成成功: {}", output);
        }
        Commands::Legacy { input, output } => {
            println!("使用传统方式生成Word文档...");
            let bills = read_data_file(input, &get_default_headers())?;
            let docx_content = generate_word_document_with_template(&bills, None)?;
            std::fs::write(output, docx_content)?;
            println!("✅ Word文档生成成功: {}", output);
        }
    }

    Ok(())
}

fn get_default_headers() -> HeadersMap<'static> {
    HeadersMap {
        merchant: "店铺名称",
        water_prev: "上期水表读数",
        water_curr: "本期水表读数",
        water_price: "水费单价",
        electricity_price: "电费单价",
        electricity_prefix: "电表",
    }
}

// 辅助函数
fn normalize(s: &str) -> String { s.trim().to_lowercase() }

fn find_indices(headers: &[String], mapping: &[(&str, &str)]) -> Result<Vec<usize>> {
    let headers_norm: Vec<String> = headers.iter().map(|h| normalize(h)).collect();
    let mut indices = Vec::with_capacity(mapping.len());

    for (want, fallback) in mapping.iter() {
        let candidates = vec![normalize(want), normalize(fallback)];
        let mut found = headers_norm.iter().position(|h| candidates.iter().any(|c| c == h));
        if found.is_none() {
            found = headers_norm.iter().position(|h| candidates.iter().any(|c| h.contains(c)));
        }
        let idx = found.with_context(|| format!("无法在表头中找到列: {}", want))?;
        indices.push(idx);
    }
    Ok(indices)
}

fn as_f64(cell: &DataType) -> f64 {
    match cell {
        DataType::Float(f) => *f,
        DataType::Int(i) => *i as f64,
        DataType::String(s) => s.trim().parse::<f64>().unwrap_or(0.0),
        _ => 0.0,
    }
}

fn read_excel_file(file_path: &str, headers_map: &HeadersMap) -> Result<Vec<MerchantBill>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path)
        .with_context(|| format!("无法打开Excel文件: {}", file_path))?;
    let sheet_name = workbook.sheet_names()[0].clone();
    let range = workbook
        .worksheet_range(&sheet_name)
        .with_context(|| format!("无法读取工作表: {}", sheet_name))??;

    let mut rows = range.rows();
    let header_row = rows.next().context("Excel中缺少表头行")?;
    let headers: Vec<String> = header_row.iter().map(|c| c.to_string()).collect();
    
    println!("调试：Excel表头: {:?}", headers);
    
    // 直接查找列索引，不使用find_indices
    let m_i = headers.iter().position(|h| h.contains("店铺名称")).context("找不到店铺名称列")?;
    let wp_i = headers.iter().position(|h| h.contains("上期水表读数")).context("找不到上期水表读数列")?;
    let wc_i = headers.iter().position(|h| h.contains("本期水表读数")).context("找不到本期水表读数列")?;
    let wprice_i = headers.iter().position(|h| h.contains("水费单价")).context("找不到水费单价列")?;
    let eprice_i = headers.iter().position(|h| h.contains("电费单价")).context("找不到电费单价列")?;

    // 找到所有电表相关的列
    let electricity_columns = find_electricity_columns(&headers, headers_map.electricity_prefix)?;

    println!("调试：Excel基础列索引 - 商家:{}, 水表上期:{}, 水表本期:{}, 水费单价:{}, 电费单价:{}", 
             m_i, wp_i, wc_i, wprice_i, eprice_i);
    println!("调试：Excel电表列: {:?}", electricity_columns);

    let mut bills = Vec::new();
    for row in rows {
        if row.is_empty() { continue; }
        let merchant_name = row.get(m_i).map(|c| c.to_string()).unwrap_or_default();
        if merchant_name.trim().is_empty() { continue; }
        
        let water_price = row.get(wprice_i).map(as_f64).unwrap_or(0.0);
        let electricity_price = row.get(eprice_i).map(as_f64).unwrap_or(0.0);
        let prev_water = row.get(wp_i).map(as_f64).unwrap_or(0.0);
        let curr_water = row.get(wc_i).map(as_f64).unwrap_or(0.0);

        let mut bill = MerchantBill::new(merchant_name, water_price, electricity_price);
        bill.set_water_readings(prev_water, curr_water);

        // 处理每个电表
        for (meter_id, (prev_col, curr_col)) in electricity_columns.iter().enumerate() {
            let prev_reading = row.get(*prev_col).map(as_f64).unwrap_or(0.0);
            let curr_reading = row.get(*curr_col).map(as_f64).unwrap_or(0.0);
            if prev_reading > 0.0 || curr_reading > 0.0 {
                bill.add_electricity_meter(format!("{}", meter_id + 1), prev_reading, curr_reading);
            }
        }

        // 设置人工费和垃圾处理费（这里使用固定值作为示例，实际应该从数据中读取）
        let labor_fee = 50.0; // 水电人工费
        let garbage_fee = 30.0; // 垃圾处理费
        bill.set_additional_fees(labor_fee, garbage_fee);

        bills.push(bill);
    }
    Ok(bills)
}

fn read_csv_file(file_path: &str, headers_map: &HeadersMap) -> Result<Vec<MerchantBill>> {
    let file = File::open(file_path)
        .with_context(|| format!("无法打开CSV文件: {}", file_path))?;
    let mut lines = BufReader::new(file).lines();
    let header_line = lines.next().transpose()?.context("CSV中缺少表头行")?;
    let headers: Vec<String> = header_line.split(',').map(|s| s.trim().to_string()).collect();

    println!("调试：找到的表头: {:?}", headers);

    // 直接查找列索引，不使用find_indices
    let m_i = headers.iter().position(|h| h.contains("店铺名称")).context("找不到店铺名称列")?;
    let wp_i = headers.iter().position(|h| h.contains("上期水表读数")).context("找不到上期水表读数列")?;
    let wc_i = headers.iter().position(|h| h.contains("本期水表读数")).context("找不到本期水表读数列")?;
    let wprice_i = headers.iter().position(|h| h.contains("水费单价")).context("找不到水费单价列")?;
    let eprice_i = headers.iter().position(|h| h.contains("电费单价")).context("找不到电费单价列")?;

    // 找到所有电表相关的列
    let electricity_columns = find_electricity_columns(&headers, headers_map.electricity_prefix)?;

    println!("调试：基础列索引 - 商家:{}, 水表上期:{}, 水表本期:{}, 水费单价:{}, 电费单价:{}", 
             m_i, wp_i, wc_i, wprice_i, eprice_i);
    println!("调试：电表列: {:?}", electricity_columns);

    let mut bills = Vec::new();
    for line in lines {
        let line = line?;
        if line.trim().is_empty() { continue; }
        let parts: Vec<&str> = line.split(',').collect();
        if parts.len() < 5 { continue; } // 确保至少有基础列
        
        let get = |i: usize| -> &str { parts.get(i).copied().unwrap_or("") };
        
        let merchant_name = get(m_i).trim().to_string();
        if merchant_name.is_empty() { continue; }
        
        let water_price = get(wprice_i).trim().parse::<f64>().unwrap_or(0.0);
        let electricity_price = get(eprice_i).trim().parse::<f64>().unwrap_or(0.0);
        let prev_water = get(wp_i).trim().parse::<f64>().unwrap_or(0.0);
        let curr_water = get(wc_i).trim().parse::<f64>().unwrap_or(0.0);

        let mut bill = MerchantBill::new(merchant_name, water_price, electricity_price);
        bill.set_water_readings(prev_water, curr_water);

        // 处理每个电表
        for (meter_id, (prev_col, curr_col)) in electricity_columns.iter().enumerate() {
            let prev_reading = get(*prev_col).trim().parse::<f64>().unwrap_or(0.0);
            let curr_reading = get(*curr_col).trim().parse::<f64>().unwrap_or(0.0);
            if prev_reading > 0.0 || curr_reading > 0.0 {
                bill.add_electricity_meter(format!("{}", meter_id + 1), prev_reading, curr_reading);
            }
        }

        // 设置人工费和垃圾处理费（这里使用固定值作为示例，实际应该从数据中读取）
        let labor_fee = 50.0; // 水电人工费
        let garbage_fee = 30.0; // 垃圾处理费
        bill.set_additional_fees(labor_fee, garbage_fee);

        bills.push(bill);
    }
    Ok(bills)
}

fn find_electricity_columns(headers: &[String], prefix: &str) -> Result<Vec<(usize, usize)>> {
    let mut columns = Vec::new();
    let headers_norm: Vec<String> = headers.iter().map(|h| normalize(h)).collect();
    
    // 查找电表列的模式：电表1上期读数、电表1本期读数、电表2上期读数、电表2本期读数...
    let mut meter_id = 1;
    loop {
        let prev_pattern = format!("{}{}上期读数", prefix, meter_id);
        let curr_pattern = format!("{}{}本期读数", prefix, meter_id);
        
        let prev_idx = headers_norm.iter().position(|h| h.contains(&normalize(&prev_pattern)));
        let curr_idx = headers_norm.iter().position(|h| h.contains(&normalize(&curr_pattern)));
        
        if prev_idx.is_some() && curr_idx.is_some() {
            columns.push((prev_idx.unwrap(), curr_idx.unwrap()));
            meter_id += 1;
        } else {
            break;
        }
    }
    
    if columns.is_empty() {
        anyhow::bail!("未找到任何电表列，请确保CSV包含'电表X上期读数'和'电表X本期读数'列");
    }
    
    Ok(columns)
}

fn read_data_file(file_path: &str, headers_map: &HeadersMap) -> Result<Vec<MerchantBill>> {
    let path = Path::new(file_path);
    let extension = path.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
    match extension.as_str() {
        "xlsx" => read_excel_file(file_path, headers_map),
        "csv" => read_csv_file(file_path, headers_map),
        _ => {
            if file_path.ends_with(".xlsx") { read_excel_file(file_path, headers_map) }
            else if file_path.ends_with(".csv") { read_csv_file(file_path, headers_map) }
            else { anyhow::bail!("不支持的文件格式: {}", extension) }
        }
    }
}

fn generate_word_document_with_template(
    merchants: &[MerchantBill],
    config_path: Option<&str>,
) -> Result<Vec<u8>, anyhow::Error> {
    // 简单的模板生成，直接使用docx-rs
    use docx_rs::*;
    
    let mut doc = Docx::new();
    
    // 添加文档标题
    doc = doc.add_paragraph(
        Paragraph::new()
            .add_run(Run::new().add_text("商家水费电费账单").size(24))
            .align(AlignmentType::Center)
    );
    
    // 为每个商家生成账单
    for (index, bill) in merchants.iter().enumerate() {
        // 商家名称
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("商家名称：{}", bill.merchant_name)).size(16).bold())
        );
        
        // 账单期间
        let now = Local::now();
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("账单期间：{}年{}月", now.year(), now.month())).size(14))
        );
        
        // 水表读数
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("水表读数").size(16).bold())
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("上期水表读数：{} 吨", bill.prev_water_reading)).size(14))
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("本期水表读数：{} 吨", bill.curr_water_reading)).size(14))
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("本月用水量：{} 吨", bill.water_usage)).size(14))
        );
        
        // 电表信息
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("电表信息（共{}个电表）", bill.electricity_meters.len())).size(16).bold())
        );
        
        for meter in &bill.electricity_meters {
            doc = doc.add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text(
                        format!("电表{}: 上期{}度, 本期{}度, 用量{}度, 费用{:.2}元", 
                            meter.meter_id, 
                            meter.prev_reading, 
                            meter.curr_reading, 
                            meter.usage, 
                            meter.amount)
                    ).size(14))
            );
        }
        
        // 用量汇总
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("用量汇总").size(16).bold())
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("本月总用电量：{} 度", bill.electricity_usage)).size(14))
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("本月总用水量：{} 吨", bill.water_usage)).size(14))
        );
        
        // 费用计算
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("费用计算").size(16).bold())
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("电费单价：{:.2} 元/度", bill.electricity_unit_price)).size(14))
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("水费单价：{:.2} 元/吨", bill.water_unit_price)).size(14))
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("电费总额：{:.2} 元", bill.electricity_amount)).size(14))
        );
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("水费总额：{:.2} 元", bill.water_amount)).size(14))
        );
        
        // 费用合计
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("费用合计：{:.2} 元", bill.total_fee)).size(16).bold().color("FF0000"))
        );
        
        // 生成时间
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(format!("生成时间：{}", Local::now().format("%Y-%m-%d %H:%M:%S"))).size(10))
                .align(AlignmentType::Right)
        );
        
        // 添加分页符（除了最后一个）
        if index < merchants.len() - 1 {
            doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));
        }
    }
    
    // 添加汇总表格
    doc = add_summary_table(doc, merchants)?;
    
    let mut buf = Vec::new();
    doc.build().pack(&mut std::io::Cursor::new(&mut buf))?;
    Ok(buf)
}

fn add_summary_table(mut doc: docx_rs::Docx, merchants: &[MerchantBill]) -> Result<docx_rs::Docx, anyhow::Error> {
    use docx_rs::*;
    
    // 添加汇总表格标题
    doc = doc.add_paragraph(
        Paragraph::new()
            .add_run(Run::new().add_text("费用汇总表").size(18).bold())
            .align(AlignmentType::Center)
    );
    
    // 创建表格
    let mut table = Table::new(vec![
        TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("店铺名称").bold())),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("水电费合计（元）").bold())),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("水电人工费").bold())),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("垃圾处理费").bold())),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("总价").bold())),
        ])
    ]);

    // 添加数据行
    for bill in merchants {
        let water_electricity_total = bill.water_amount + bill.electricity_amount;
        table = table.add_row(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&bill.merchant_name))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", water_electricity_total)))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", bill.water_electricity_labor_fee)))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", bill.garbage_disposal_fee)))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", bill.total_fee)))),
        ]));
    }

    // 添加合计行
    let total_water_electricity: f64 = merchants.iter().map(|b| b.water_amount + b.electricity_amount).sum();
    let total_labor_fee: f64 = merchants.iter().map(|b| b.water_electricity_labor_fee).sum();
    let total_garbage_fee: f64 = merchants.iter().map(|b| b.garbage_disposal_fee).sum();
    let grand_total: f64 = merchants.iter().map(|b| b.total_fee).sum();

    table = table.add_row(TableRow::new(vec![
        TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("合计").bold())),
        TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", total_water_electricity)).bold())),
        TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", total_labor_fee)).bold())),
        TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", total_garbage_fee)).bold())),
        TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", grand_total)).bold())),
    ]));

    doc = doc.add_table(table);
    Ok(doc)
}