use anyhow::{Context, Result};
use crate::MerchantBill;
use chrono::{Datelike, Local};
use docx_rs::*;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;

#[derive(Debug, Deserialize, Clone)]
pub struct TemplateConfig {
    pub document_title: String,
    pub title_font_size: usize,
    pub title_alignment: String,
    pub section_font_size: usize,
    pub timestamp_font_size: usize,
    pub merchant_template: MerchantTemplate,
    pub summary_template: SummaryTemplate,
    pub output_format: String,
    pub default_output_name: String,
    pub individual_bills: bool,
    pub summary_table: bool,
}

#[derive(Debug, Deserialize, Clone)]
pub struct MerchantTemplate {
    pub sections: Vec<Section>,
}

#[derive(Debug, Deserialize, Clone)]
pub struct SummaryTemplate {
    pub sections: Vec<Section>,
}

#[derive(Debug, Deserialize, Clone)]
pub struct Section {
    pub name: String,
    pub r#type: String,
    pub content: Option<String>,
    pub title: Option<String>,
    pub items: Option<Vec<String>>,
    pub font_size: Option<usize>,
    pub bold: Option<bool>,
    pub color: Option<String>,
    pub alignment: Option<String>,
}

impl TemplateConfig {
    pub fn load_from_file(path: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let content = std::fs::read_to_string(path)?;
        let config: TemplateConfig = serde_json::from_str(&content)?;
        Ok(config)
    }

    pub fn load_default() -> Self {
        serde_json::from_str(include_str!("../config/template_config.json")).unwrap()
    }
}

pub struct DocumentGenerator {
    config: TemplateConfig,
}

impl DocumentGenerator {
    pub fn new(config: TemplateConfig) -> Self {
        Self { config }
    }

    // 生成单个商家账单
    pub fn generate_merchant_bill(&self, bill: &MerchantBill) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut doc = Docx::new();
        
        for section in &self.config.merchant_template.sections {
            match section.r#type.as_str() {
                "title" => {
                    if let Some(content) = &section.content {
                        let title_content = self.replace_placeholders(content, bill);
                        doc = doc.add_paragraph(
                            Paragraph::new()
                                .add_run(Run::new().add_text(&title_content).size(self.config.title_font_size * 2))
                                .align(if self.config.title_alignment == "center" { AlignmentType::Center } else { AlignmentType::Left })
                        );
                    }
                }
                "text" => {
                    if let Some(content) = &section.content {
                        let text_content = self.replace_placeholders(content, bill);
                        let mut run = Run::new().add_text(&text_content).size(section.font_size.unwrap_or(self.config.section_font_size));
                        
                        if section.bold.unwrap_or(false) {
                            run = run.bold();
                        }
                        
                        if let Some(color) = &section.color {
                            run = run.color(color);
                        }
                        
                        let mut paragraph = Paragraph::new().add_run(run);
                        
                        if let Some(alignment) = &section.alignment {
                            paragraph = paragraph.align(match alignment.as_str() {
                                "center" => AlignmentType::Center,
                                "right" => AlignmentType::Right,
                                _ => AlignmentType::Left,
                            });
                        }
                        
                        doc = doc.add_paragraph(paragraph);
                    }
                }
                "section" => {
                    if let Some(title) = &section.title {
                        // 添加小标题
                        doc = doc.add_paragraph(
                            Paragraph::new()
                                .add_run(Run::new().add_text(title).bold().size((self.config.section_font_size + 4) * 2))
                        );
                    }

                    if let Some(items) = &section.items {
                        for item in items {
                            let item_content = self.replace_placeholders(item, bill);
                            doc = doc.add_paragraph(
                                Paragraph::new()
                                    .add_run(Run::new().add_text(&item_content).size(self.config.section_font_size * 2))
                            );
                        }
                    }
                }
                "timestamp" => {
                    if let Some(format) = &section.content {
                        let datetime = Local::now();
                        let timestamp_content = format
                            .replace("{datetime}", &datetime.format("%Y-%m-%d %H:%M:%S").to_string());
                        
                        let mut paragraph = Paragraph::new().add_run(
                            Run::new().add_text(&timestamp_content).size(self.config.timestamp_font_size)
                        );
                        
                        if let Some(alignment) = &section.alignment {
                            paragraph = paragraph.align(match alignment.as_str() {
                                "center" => AlignmentType::Center,
                                "right" => AlignmentType::Right,
                                _ => AlignmentType::Left,
                            });
                        }
                        
                        doc = doc.add_paragraph(paragraph);
                    }
                }
                _ => {}
            }
        }
        
        // 添加分页符（除了最后一个）
        doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));
        
        let mut buf = Vec::new();
        doc.build().pack()?.write(&mut buf)?;
        Ok(buf)
    }

    // 生成汇总表格（可选）
    pub fn generate_summary_table(&self, bills: &[MerchantBill]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut doc = Docx::new();
        
        for section in &self.config.summary_template.sections {
            match section.r#type.as_str() {
                "title" => {
                    if let Some(content) = &section.content {
                        doc = doc.add_paragraph(
                            Paragraph::new()
                                .add_run(Run::new().add_text(content).size(self.config.title_font_size * 2))
                                .align(AlignmentType::Center)
                        );
                    }
                }
                "text" => {
                    if let Some(content) = &section.content {
                        let text_content = self.replace_placeholders(content, &bills[0]);
                        doc = doc.add_paragraph(
                            Paragraph::new()
                                .add_run(Run::new().add_text(&text_content).size(self.config.section_font_size))
                        );
                    }
                }
                "table" => {
                    doc = self.create_summary_table(doc, bills)?;
                }
                "timestamp" => {
                    if let Some(format) = &section.content {
                        let datetime = Local::now();
                        let timestamp_content = format
                            .replace("{datetime}", &datetime.format("%Y-%m-%d %H:%M:%S").to_string());
                        
                        doc = doc.add_paragraph(
                            Paragraph::new()
                                .add_run(Run::new().add_text(&timestamp_content).size(self.config.timestamp_font_size))
                                .align(AlignmentType::Right)
                        );
                    }
                }
                _ => {}
            }
        }
        
        let mut buf = Vec::new();
        doc.build().pack()?.write(&mut buf)?;
        Ok(buf)
    }

    // 生成完整文档（包含所有商家账单）
    pub fn generate_complete_document(&self, bills: &[MerchantBill]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut doc = Docx::new();
        
        // 添加文档标题
        doc = doc.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(&self.config.document_title).size(self.config.title_font_size * 2))
                .align(AlignmentType::Center)
        );
        
        // 为每个商家生成账单
        for (index, bill) in bills.iter().enumerate() {
            // 添加商家账单
            for section in &self.config.merchant_template.sections {
                match section.r#type.as_str() {
                    "title" => {
                        if let Some(content) = &section.content {
                            let title_content = self.replace_placeholders(content, bill);
                            doc = doc.add_paragraph(
                                Paragraph::new()
                                    .add_run(Run::new().add_text(&title_content).size(self.config.title_font_size * 2))
                                    .align(AlignmentType::Center)
                            );
                        }
                    }
                    "text" => {
                        if let Some(content) = &section.content {
                            let text_content = self.replace_placeholders(content, bill);
                            let mut run = Run::new().add_text(&text_content).size(section.font_size.unwrap_or(self.config.section_font_size));
                            
                            if section.bold.unwrap_or(false) {
                                run = run.bold();
                            }
                            
                            if let Some(color) = &section.color {
                                run = run.color(color);
                            }
                            
                            let mut paragraph = Paragraph::new().add_run(run);
                            
                            if let Some(alignment) = &section.alignment {
                                paragraph = paragraph.align(match alignment.as_str() {
                                    "center" => AlignmentType::Center,
                                    "right" => AlignmentType::Right,
                                    _ => AlignmentType::Left,
                                });
                            }
                            
                            doc = doc.add_paragraph(paragraph);
                        }
                    }
                    "section" => {
                        if let Some(title) = &section.title {
                            doc = doc.add_paragraph(
                                Paragraph::new()
                                    .add_run(Run::new().add_text(title).bold().size((self.config.section_font_size + 4) * 2))
                            );
                        }

                        if let Some(items) = &section.items {
                            for item in items {
                                let item_content = self.replace_placeholders(item, bill);
                                doc = doc.add_paragraph(
                                    Paragraph::new()
                                        .add_run(Run::new().add_text(&item_content).size(self.config.section_font_size * 2))
                                );
                            }
                        }
                    }
                    "timestamp" => {
                        if let Some(format) = &section.content {
                            let datetime = Local::now();
                            let timestamp_content = format
                                .replace("{datetime}", &datetime.format("%Y-%m-%d %H:%M:%S").to_string());
                            
                            let mut paragraph = Paragraph::new().add_run(
                                Run::new().add_text(&timestamp_content).size(self.config.timestamp_font_size)
                            );
                            
                            if let Some(alignment) = &section.alignment {
                                paragraph = paragraph.align(match alignment.as_str() {
                                    "center" => AlignmentType::Center,
                                    "right" => AlignmentType::Right,
                                    _ => AlignmentType::Left,
                                });
                            }
                            
                            doc = doc.add_paragraph(paragraph);
                        }
                    }
                    _ => {}
                }
            }
            
            // 添加分页符（除了最后一个）
            if index < bills.len() - 1 {
                doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));
            }
        }
        
        let mut buf = Vec::new();
        doc.build().pack()?.write(&mut buf)?;
        Ok(buf)
    }

    fn replace_placeholders(&self, text: &str, bill: &MerchantBill) -> String {
        let datetime = Local::now();
        let mut result = text.to_string();
        
        // 替换商家信息
        result = result.replace("{merchant_name}", &bill.merchant_name);
        result = result.replace("{year}", &datetime.year().to_string());
        result = result.replace("{month}", &datetime.month().to_string());
        
        // 替换表计读数
        result = result.replace("{prev_electric_reading}", &bill.prev_electric_reading.to_string());
        result = result.replace("{curr_electric_reading}", &bill.curr_electric_reading.to_string());
        result = result.replace("{prev_water_reading}", &bill.prev_water_reading.to_string());
        result = result.replace("{curr_water_reading}", &bill.curr_water_reading.to_string());
        
        // 替换用量计算
        result = result.replace("{electricity_usage}", &bill.electricity_usage.to_string());
        result = result.replace("{water_usage}", &bill.water_usage.to_string());
        
        // 替换费用计算
        result = result.replace("{electricity_unit_price}", &format!("{:.2}", bill.electricity_unit_price));
        result = result.replace("{water_unit_price}", &format!("{:.2}", bill.water_unit_price));
        result = result.replace("{electricity_amount}", &format!("{:.2}", bill.electricity_amount));
        result = result.replace("{water_amount}", &format!("{:.2}", bill.water_amount));
        result = result.replace("{total_amount}", &format!("{:.2}", bill.total_fee));
        
        result
    }

    fn create_summary_table(&self, mut doc: Docx, bills: &[MerchantBill]) -> Result<Docx, Box<dyn std::error::Error>> {
        // 创建表格，设置列宽和表头
        let mut table = Table::new(vec![
            TableRow::new(vec![
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text("序号").bold().size(40)))
                    .width(1200, WidthType::Dxa),
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text("商家名称").bold().size(40)))
                    .width(4000, WidthType::Dxa),
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text("水费(元)").bold().size(40)))
                    .width(2400, WidthType::Dxa),
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text("电费(元)").bold().size(40)))
                    .width(2400, WidthType::Dxa),
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text("合计(元)").bold().size(40)))
                    .width(2400, WidthType::Dxa),
            ])
            .height(800, HeightRule::AtLeast)
        ])
        .width(12400, WidthType::Dxa);

        // 添加数据行
        for (index, bill) in bills.iter().enumerate() {
            table = table.add_row(TableRow::new(vec![
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text((index + 1).to_string()).size(36)))
                    .width(1200, WidthType::Dxa),
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&bill.merchant_name).size(36)))
                    .width(4000, WidthType::Dxa),
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", bill.water_amount)).size(36)))
                    .width(2400, WidthType::Dxa),
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", bill.electricity_amount)).size(36)))
                    .width(2400, WidthType::Dxa),
                TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", bill.total_fee)).size(36)))
                    .width(2400, WidthType::Dxa),
            ])
            .height(700, HeightRule::AtLeast));
        }

        // 添加合计行
        let total_water: f64 = bills.iter().map(|b| b.water_amount).sum();
        let total_electricity: f64 = bills.iter().map(|b| b.electricity_amount).sum();
        let grand_total: f64 = bills.iter().map(|b| b.total_fee).sum();

        table = table.add_row(TableRow::new(vec![
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("合计").bold().size(40)))
                .width(1200, WidthType::Dxa),
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("").bold().size(40)))
                .width(4000, WidthType::Dxa),
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", total_water)).bold().size(40)))
                .width(2400, WidthType::Dxa),
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", total_electricity)).bold().size(40)))
                .width(2400, WidthType::Dxa),
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!("{:.2}", grand_total)).bold().size(40)))
                .width(2400, WidthType::Dxa),
        ])
        .height(800, HeightRule::AtLeast));

        doc = doc.add_table(table);
        Ok(doc)
    }
}
