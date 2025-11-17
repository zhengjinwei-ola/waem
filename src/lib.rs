use anyhow::{Context, Result};
use calamine::{open_workbook, DataType, Reader, Xlsx};
use chrono::{Local, Datelike};
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::Path;

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
    pub shop_code: String, // 铺面编号（字符串）
    pub water_unit_price: f64,
    pub electricity_unit_price: f64,
    pub prev_water_reading: f64,
    pub curr_water_reading: f64,
    pub water_usage: f64,
    pub water_amount: f64,
    pub electricity_meters: Vec<ElectricityMeter>,
    pub electricity_usage: f64,
    pub electricity_amount: f64,
    pub water_electricity_labor_fee: f64,  // 水电人工费
    pub garbage_disposal_fee: f64,         // 垃圾处理费
    pub meter_reader: Option<String>,      // 抄表人（可选，由Web表单传入）
    pub meter_date: Option<String>,        // 抄表日期（可选，由Web表单传入）
    pub total_fee: f64,
    pub month: String,
}

#[derive(Debug)]
pub struct BillTemplate {
    pub month: String,
    pub year: String,
    pub merchants: Vec<MerchantBill>,
    pub total_water_usage: f64,
    pub total_electric_usage: f64,
    pub total_water_amount: f64,
    pub total_electric_amount: f64,
    pub grand_total: f64,
}

impl MerchantBill {
    pub fn new(merchant_name: String, water_unit_price: f64, electricity_unit_price: f64) -> Self {
        Self {
            merchant_name,
            shop_code: String::new(),
            water_unit_price,
            electricity_unit_price,
            prev_water_reading: 0.0,
            curr_water_reading: 0.0,
            water_usage: 0.0,
            water_amount: 0.0,
            electricity_meters: Vec::new(),
            electricity_usage: 0.0,
            electricity_amount: 0.0,
            water_electricity_labor_fee: 0.0,  // 水电人工费
            garbage_disposal_fee: 0.0,         // 垃圾处理费
            meter_reader: None,
            meter_date: None,
            total_fee: 0.0,
            month: Local::now().format("%Y年%m月").to_string(),
        }
    }

    pub fn set_shop_code(&mut self, code: String) { self.shop_code = code; }
    pub fn set_meter_info(&mut self, reader: Option<String>, date: Option<String>) {
        self.meter_reader = reader;
        self.meter_date = date;
    }

    pub fn set_water_readings(&mut self, prev: f64, curr: f64) {
        self.prev_water_reading = prev;
        self.curr_water_reading = curr;
        self.water_usage = (curr - prev).max(0.0);
        // 水费金额四舍五入到"元"（整数）
        self.water_amount = (self.water_usage * self.water_unit_price).round();
        self.update_totals();
    }

    pub fn add_electricity_meter(&mut self, meter_id: String, prev: f64, curr: f64) {
        let usage = (curr - prev).max(0.0);
        // 行内展示用的单表金额（四舍五入到元，仅展示用）
        let amount = (usage * self.electricity_unit_price).round();
        self.electricity_meters.push(ElectricityMeter {
            meter_id,
            prev_reading: prev,
            curr_reading: curr,
            usage,
            amount,
        });
        self.update_totals();
    }

    pub fn update_totals(&mut self) {
        // 总用电量
        self.electricity_usage = self.electricity_meters.iter().map(|m| m.usage).sum();
        // 电费按规则：先合计总用电量，再乘单价，最后四舍五入到元
        self.electricity_amount = (self.electricity_usage * self.electricity_unit_price).round();
        // 水费金额已在设置时四舍五入到元
        // 总费用根据电费总额(总用量*单价后四舍五入)、水费(四舍五入后)与其他费用直接相加
        self.total_fee = self.water_amount + self.electricity_amount + self.water_electricity_labor_fee + self.garbage_disposal_fee;
    }

    pub fn get_electricity_details(&self) -> String {
        if self.electricity_meters.is_empty() {
            return "无电表数据".to_string();
        }
        
        let details: Vec<String> = self.electricity_meters.iter().map(|meter| {
            format!("电表{}: 上期{}度, 本期{}度, 用量{}度, 费用{:.2}元", 
                meter.meter_id, meter.prev_reading, meter.curr_reading, meter.usage, meter.amount)
        }).collect();
        
        details.join("\n")
    }
}

impl BillTemplate {
    pub fn new(month: String, year: String) -> Self {
        Self {
            month,
            year,
            merchants: Vec::new(),
            total_water_usage: 0.0,
            total_electric_usage: 0.0,
            total_water_amount: 0.0,
            total_electric_amount: 0.0,
            grand_total: 0.0,
        }
    }

    pub fn add_merchant(&mut self, merchant: MerchantBill) {
        self.total_water_usage += merchant.water_usage;
        self.total_electric_usage += merchant.electricity_usage;
        self.total_water_amount += merchant.water_amount;
        self.total_electric_amount += merchant.electricity_amount;
        self.grand_total += merchant.total_fee;
        self.merchants.push(merchant);
    }
}

#[derive(Clone)]
pub struct HeadersMap<'a> {
    pub merchant: &'a str,
    pub prev_e: &'a str,
    pub curr_e: &'a str,
    pub prev_w: &'a str,
    pub curr_w: &'a str,
    pub w_price: &'a str,
    pub e_price: &'a str,
    pub electricity_price: &'a str,
    pub electricity_prefix: &'a str,
    pub water_electricity_labor_fee: &'a str,  // 水电人工费
    pub garbage_disposal_fee: &'a str,         // 垃圾处理费
}

// 已不再使用的映射帮助方法移除，避免未使用告警

fn normalize(s: &str) -> String { s.trim().to_lowercase() }

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

// 已不再使用的函数移除，避免未使用告警

fn as_f64(cell: &DataType) -> f64 {
    match cell {
        DataType::Float(f) => *f,
        DataType::Int(i) => *i as f64,
        DataType::String(s) => s.trim().parse::<f64>().unwrap_or(0.0),
        _ => 0.0,
    }
}

pub struct GenerateOptions {
    pub custom_title: Option<String>,
    pub per_page: usize,
}

pub fn generate_word_document_with_template(
    merchants: &[MerchantBill],
    options: Option<GenerateOptions>,
) -> Result<Vec<u8>, anyhow::Error> {
    // 生成专业的抄表计费通知单格式（表格版）
    use docx_rs::*;
    
    let mut doc = Docx::new();
    
    let per_page = options.as_ref().map(|o| o.per_page).unwrap_or(1);
    // 为每个商家生成通知单
    for (index, bill) in merchants.iter().enumerate() {
        let now = Local::now();
        let year = now.year();
        let month = now.month();
        let day = now.day();
        
        // 标题：自定义或默认 "yyyy年MM月抄表计费通知单"
        let title = options
            .as_ref()
            .and_then(|o| o.custom_title.clone())
            .unwrap_or_else(|| format!("{}年{:02}月抄表计费通知单", year, month));
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(&title).bold().size(22))
                .align(AlignmentType::Center)
        );
        
        // 编号和基本信息行（编号使用CSV的铺面编号；抄表人/日期来自页面输入）
        let meter_reader = bill.meter_reader.clone().unwrap_or_else(|| "".to_string());
        let meter_date = bill.meter_date.clone().unwrap_or_else(|| format!("{}年{:02}月{:02}日", year, month, day));
        let info_text = format!("编号：\t{}\t姓名\t{}\t抄表人：\t{}\t抄表日期：{}", 
            bill.shop_code, bill.merchant_name, meter_reader, meter_date);
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(&info_text).size(11))
        );
        
        // 空行
        doc = doc.add_paragraph(Paragraph::new());
        
        // 创建费用明细表格
        let mut table_rows = vec![
            TableRow::new(vec![
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("项目").bold().size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("上月表底").bold().size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("本月抄表数").bold().size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("实用度数").bold().size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("公共分摊").bold().size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("单价（元）").bold().size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("金额").bold().size(12)).align(AlignmentType::Center)),
            ]),
        ];
        
        // 为每个电表生成行；若电表>1，仅在最后一行显示合并后的“金额”
        let meters_len = bill.electricity_meters.len();
        for (meter_idx, meter) in bill.electricity_meters.iter().enumerate() {
            let meter_name = if meters_len == 1 {
                "电表".to_string()
            } else {
                format!("电表{}", meter_idx + 1)
            };

            // 单价与金额列：若>1电表，对这两列做纵向合并（类似Excel合并单元格）
            // 合并策略：
            // - 单价列：首行显示单价并 vMerge Restart，其余行 vMerge Continue
            // - 金额列：首行显示合并后的电费总额并 vMerge Restart，其余行 vMerge Continue
            // 若仅1个电表，则正常显示，无合并

            // 构造单价列单元格（第6列）
            let unit_price_cell = if meters_len > 1 {
                if meter_idx == 0 {
                    TableCell::new()
                        .vertical_merge(VMergeType::Restart)
                        .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.2}", bill.electricity_unit_price)).size(12)).align(AlignmentType::Center))
                } else {
                    TableCell::new()
                        .vertical_merge(VMergeType::Continue)
                }
            } else {
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.2}", bill.electricity_unit_price)).size(12)).align(AlignmentType::Center))
            };

            // 构造金额列单元格（第7列）
            let amount_cell = if meters_len > 1 {
                if meter_idx == 0 {
                    TableCell::new()
                        .vertical_merge(VMergeType::Restart)
                        .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", bill.electricity_amount)).size(12)).align(AlignmentType::Center))
                } else {
                    TableCell::new()
                        .vertical_merge(VMergeType::Continue)
                }
            } else {
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", bill.electricity_amount)).size(12)).align(AlignmentType::Center))
            };

            table_rows.push(TableRow::new(vec![
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&meter_name).size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", meter.prev_reading)).size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", meter.curr_reading)).size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", meter.usage)).size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
                unit_price_cell,
                amount_cell,
            ]));
        }
        
        // 如果没有电表，添加一个空行
        if bill.electricity_meters.is_empty() {
            table_rows.push(TableRow::new(vec![
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("电表").size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("0").size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("0").size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("0").size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.2}", bill.electricity_unit_price)).size(12)).align(AlignmentType::Center)),
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("0").size(12)).align(AlignmentType::Center)),
            ]));
        }
        
        // 添加水费行（去掉“损耗/实用”子行，仅保留单价与金额）
        table_rows.push(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("水费").size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", bill.prev_water_reading)).size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", bill.curr_water_reading)).size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", bill.water_usage)).size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.3}", bill.water_unit_price)).size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.0}", bill.water_amount)).size(12)).align(AlignmentType::Center)),
        ]));

        table_rows.push(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("水电人工费").size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.2}", bill.water_electricity_labor_fee)).size(12)).align(AlignmentType::Center))
        ]));

        table_rows.push(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("垃圾处理费").size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{:.2}", bill.garbage_disposal_fee)).size(12)).align(AlignmentType::Center))
        ]));

        // 添加滞纳金行（占位，金额为0）
        table_rows.push(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("滞纳金").size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("0.00").size(12)).align(AlignmentType::Center))
        ]));

        // 添加广告费行（占位，金额为0）
        table_rows.push(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("广告费").size(12)).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("")).align(AlignmentType::Center)),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("0.00").size(12)).align(AlignmentType::Center))
        ]));

        // 合计行（整行合并，先大写后小写，独占一行）
        let total_val = bill.total_fee;
        table_rows.push(TableRow::new(vec![
            // 第一列：项目名称（"合计"）
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("合计").bold().size(12)).align(AlignmentType::Center)),
            // 第二列到第七列合并：显示大写和小写金额
            TableCell::new()
                .grid_span(6)
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("大写：{}    小写：{:.2}", rmb_upper(total_val), total_val)).bold().size(12)).align(AlignmentType::Center))
        ]));

        let table = Table::new(table_rows);
        
        // 添加表格到文档
        doc = doc.add_table(table);
        
        // 已合并其他费用与合计到主表，不再添加第二个表格或表外合计
        
        // 空行
        doc = doc.add_paragraph(Paragraph::new());
        
        // 说明文字
        let notice_text = "1、此单可对账不做凭证；\n\n2、每月5日前为收费时间，超期按5%收滞纳金或停电；\n\n3、以上费用如有不明或差\n请到管理处核对。";
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(notice_text).size(9))
        );
        
        // 表格之间的分隔符，以及按每页数量分页
        if index < merchants.len() - 1 {
            // 分隔线一行
            doc = doc.add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("=".repeat(40)))
            );

            // 页面分隔：每页显示 per_page 个表格
            if per_page != 0 && ((index + 1) % per_page == 0) {
                doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));
            }
        }
    }

    // 汇总表之前添加分页符，使其单独成页
    doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));

    // 添加汇总表格
    doc = add_summary_table(doc, merchants)?;
    
    // 生成文档
    let mut buf = Vec::new();
    doc.build().pack(&mut std::io::Cursor::new(&mut buf))?;
    Ok(buf)
}

pub fn read_excel_file(file_path: &str, headers_map: &HeadersMap) -> Result<Vec<MerchantBill>> {
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
    let code_i = headers.iter().position(|h| h.contains("铺面编号")).context("找不到铺面编号列")?;
    let m_i = headers.iter().position(|h| h.contains("店铺名称")).context("找不到店铺名称列")?;
    // 新排序：优先电表1，然后水表，上到下
    let e1p_i = headers.iter().position(|h| h.contains("电表1上期读数")).context("找不到电表1上期读数列")?;
    let e1c_i = headers.iter().position(|h| h.contains("电表1本期读数")).context("找不到电表1本期读数列")?;
    let wp_i = headers.iter().position(|h| h.contains("上期水表读数")).context("找不到上期水表读数列")?;
    let wc_i = headers.iter().position(|h| h.contains("本期水表读数")).context("找不到本期水表读数列")?;
    let wprice_i = headers.iter().position(|h| h.contains("水费单价")).context("找不到水费单价列")?;
    let eprice_i = headers.iter().position(|h| h.contains("电费单价")).context("找不到电费单价列")?;

    // 找到水电人工费和垃圾处理费列
    let labor_fee_i = headers.iter().position(|h| h.contains("水电人工费")).context("找不到水电人工费列")?;
    let garbage_fee_i = headers.iter().position(|h| h.contains("垃圾处理费")).context("找不到垃圾处理费列")?;

    // 找到所有电表相关的列（包含已知的电表1）
    let mut electricity_columns = find_electricity_columns(&headers, headers_map.electricity_prefix)?;
    // 确保电表1优先（若已存在则不重复）
    if !electricity_columns.iter().any(|(p,c)| *p==e1p_i && *c==e1c_i) {
        electricity_columns.insert(0, (e1p_i, e1c_i));
    }

    println!("调试：Excel基础列索引 - 商家:{}, 水表上期:{}, 水表本期:{}, 水费单价:{}, 电费单价:{}, 水电人工费:{}, 垃圾处理费:{}", 
             m_i, wp_i, wc_i, wprice_i, eprice_i, labor_fee_i, garbage_fee_i);
    println!("调试：Excel电表列: {:?}", electricity_columns);

    let mut bills = Vec::new();
    for row in rows {
        if row.is_empty() { continue; }
        let merchant_name = row.get(m_i).map(|c| c.to_string()).unwrap_or_default();
        let shop_code = row.get(code_i).map(|c| c.to_string()).unwrap_or_default();
        if merchant_name.trim().is_empty() { continue; }
        
        let water_price = row.get(wprice_i).map(as_f64).unwrap_or(0.0);
        let electricity_price = row.get(eprice_i).map(as_f64).unwrap_or(0.0);
        let prev_water = row.get(wp_i).map(as_f64).unwrap_or(0.0);
        let curr_water = row.get(wc_i).map(as_f64).unwrap_or(0.0);

        let mut bill = MerchantBill::new(merchant_name, water_price, electricity_price);
        bill.set_water_readings(prev_water, curr_water);
        bill.set_shop_code(shop_code);

        // 处理每个电表
        for (meter_id, (prev_col, curr_col)) in electricity_columns.iter().enumerate() {
            let prev_reading = row.get(*prev_col).map(as_f64).unwrap_or(0.0);
            let curr_reading = row.get(*curr_col).map(as_f64).unwrap_or(0.0);
            if prev_reading > 0.0 || curr_reading > 0.0 {
                bill.add_electricity_meter(format!("{}", meter_id + 1), prev_reading, curr_reading);
            }
        }

        // 从Excel读取水电人工费和垃圾处理费
        let labor_fee = row.get(labor_fee_i).map(as_f64).unwrap_or(0.0);
        let garbage_fee = row.get(garbage_fee_i).map(as_f64).unwrap_or(0.0);
        bill.water_electricity_labor_fee = labor_fee;
        bill.garbage_disposal_fee = garbage_fee;
        bill.update_totals();

        bills.push(bill);
    }
    Ok(bills)
}

pub fn read_csv_file(file_path: &str, headers_map: &HeadersMap) -> Result<Vec<MerchantBill>> {
    let file = File::open(file_path)
        .with_context(|| format!("无法打开CSV文件: {}", file_path))?;
    let mut lines = BufReader::new(file).lines();
    let header_line = lines.next().transpose()?.context("CSV中缺少表头行")?;
    let headers: Vec<String> = header_line.split(',').map(|s| s.trim().to_string()).collect();

    println!("调试：找到的表头: {:?}", headers);

    // 直接查找列索引，不使用find_indices
    let code_i = headers.iter().position(|h| h.contains("铺面编号")).context("找不到铺面编号列")?;
    let m_i = headers.iter().position(|h| h.contains("店铺名称")).context("找不到店铺名称列")?;
    let e1p_i = headers.iter().position(|h| h.contains("电表1上期读数")).context("找不到电表1上期读数列")?;
    let e1c_i = headers.iter().position(|h| h.contains("电表1本期读数")).context("找不到电表1本期读数列")?;
    let wp_i = headers.iter().position(|h| h.contains("上期水表读数")).context("找不到上期水表读数列")?;
    let wc_i = headers.iter().position(|h| h.contains("本期水表读数")).context("找不到本期水表读数列")?;
    let wprice_i = headers.iter().position(|h| h.contains("水费单价")).context("找不到水费单价列")?;
    let eprice_i = headers.iter().position(|h| h.contains("电费单价")).context("找不到电费单价列")?;
    
    // 找到水电人工费和垃圾处理费列
    let labor_fee_i = headers.iter().position(|h| h.contains("水电人工费")).context("找不到水电人工费列")?;
    let garbage_fee_i = headers.iter().position(|h| h.contains("垃圾处理费")).context("找不到垃圾处理费列")?;

    let mut electricity_columns = find_electricity_columns(&headers, headers_map.electricity_prefix)?;
    if !electricity_columns.iter().any(|(p,c)| *p==e1p_i && *c==e1c_i) {
        electricity_columns.insert(0, (e1p_i, e1c_i));
    }

    println!("调试：基础列索引 - 商家:{}, 水表上期:{}, 水表本期:{}, 水费单价:{}, 电费单价:{}, 水电人工费:{}, 垃圾处理费:{}", 
             m_i, wp_i, wc_i, wprice_i, eprice_i, labor_fee_i, garbage_fee_i);
    println!("调试：电表列: {:?}", electricity_columns);

    let mut bills = Vec::new();
    for line in lines {
        let line = line?;
        if line.trim().is_empty() { continue; }
        let parts: Vec<&str> = line.split(',').collect();
        if parts.len() < 5 { continue; } // 确保至少有基础列
        
        let get = |i: usize| -> &str { parts.get(i).copied().unwrap_or("") };
        
        let merchant_name = get(m_i).trim().to_string();
        let shop_code = get(code_i).trim().to_string();
        if merchant_name.is_empty() { continue; }
        
        let water_price = get(wprice_i).trim().parse::<f64>().unwrap_or(0.0);
        let electricity_price = get(eprice_i).trim().parse::<f64>().unwrap_or(0.0);
        let prev_water = get(wp_i).trim().parse::<f64>().unwrap_or(0.0);
        let curr_water = get(wc_i).trim().parse::<f64>().unwrap_or(0.0);

        let mut bill = MerchantBill::new(merchant_name, water_price, electricity_price);
        bill.set_water_readings(prev_water, curr_water);
        bill.set_shop_code(shop_code);

        // 处理每个电表
        for (meter_id, (prev_col, curr_col)) in electricity_columns.iter().enumerate() {
            let prev_reading = get(*prev_col).trim().parse::<f64>().unwrap_or(0.0);
            let curr_reading = get(*curr_col).trim().parse::<f64>().unwrap_or(0.0);
            if prev_reading > 0.0 || curr_reading > 0.0 {
                bill.add_electricity_meter(format!("{}", meter_id + 1), prev_reading, curr_reading);
            }
        }

        // 从CSV读取水电人工费和垃圾处理费
        let labor_fee = get(labor_fee_i).trim().parse::<f64>().unwrap_or(0.0);
        let garbage_fee = get(garbage_fee_i).trim().parse::<f64>().unwrap_or(0.0);
        bill.water_electricity_labor_fee = labor_fee;
        bill.garbage_disposal_fee = garbage_fee;
        bill.update_totals();

        bills.push(bill);
    }
    Ok(bills)
}

pub fn read_data_file(file_path: &str, headers_map: &HeadersMap) -> Result<Vec<MerchantBill>> {
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

// 将数值金额转换为中文大写人民币（元到分）
fn rmb_upper(amount: f64) -> String {
    // 四舍五入到分
    let cents = (amount * 100.0).round() as i64;
    if cents == 0 {
        return "零元整".to_string();
    }

    let digits = ["零","壹","贰","叁","肆","伍","陆","柒","捌","玖"]; 
    let units = ["分","角","元","拾","佰","仟","万","拾","佰","仟","亿","拾","佰","仟","万"]; // 足够长

    let mut num = cents;
    let mut parts: Vec<String> = Vec::new();
    let mut unit_idx = 0usize;
    let mut last_zero = false;

    while num > 0 && unit_idx < units.len() {
        let d = (num % 10) as usize;
        let unit = units[unit_idx];
        if d == 0 {
            if (unit == "元" || unit == "万" || unit == "亿") && !parts.iter().any(|p| p.contains(unit)) {
                parts.push(unit.to_string());
            }
            if !last_zero { parts.push("零".to_string()); }
            last_zero = true;
        } else {
            let mut seg = String::new();
            seg.push_str(units[unit_idx]);
            seg.insert_str(0, digits[d]);
            parts.push(seg);
            last_zero = false;
        }
        num /= 10;
        unit_idx += 1;
    }

    parts.reverse();
    let mut s = parts.join("");
    // 清理多余的零
    while s.contains("零零") { s = s.replace("零零", "零"); }
    s = s.replace("零亿", "亿").replace("零万", "万").replace("零元", "元");
    if s.ends_with("零") { s.pop(); }
    if !s.contains("角") && !s.contains("分") { s.push_str("整"); }
    s
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
