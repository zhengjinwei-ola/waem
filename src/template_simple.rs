use crate::MerchantBill;
use chrono::{Datelike, Local};
use docx_rs::*;
use serde::Deserialize;
use std::clone::Clone;

#[derive(Debug, Deserialize, Clone)]
pub struct TemplateConfig {
    pub document_title: String,
    pub title_font_size: usize,
    pub title_alignment: String,
    pub section_font_size: usize,
    pub timestamp_font_size: usize,
    pub merchant_template: MerchantTemplate,
    pub output_format: String,
    pub default_output_name: String,
    pub individual_bills: bool,
}

#[derive(Debug, Deserialize, Clone)]
pub struct MerchantTemplate {
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
                            
                            doc = doc.add_paragraph(Paragraph::new().add_run(run));
                        }
                    }
                    "section" => {
                        if let Some(title) = &section.title {
                            doc = doc.add_paragraph(
                                Paragraph::new()
                                    .add_run(Run::new().add_text(title).bold().size(self.config.section_font_size + 2))
                            );
                        }
                        
                        if let Some(items) = &section.items {
                            for item in items {
                                let item_content = self.replace_placeholders(item, bill);
                                doc = doc.add_paragraph(
                                    Paragraph::new()
                                        .add_run(Run::new().add_text(&item_content).size(self.config.section_font_size))
                                );
                            }
                        }
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
            
            // 添加分页符（除了最后一个）
            if index < bills.len() - 1 {
                doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));
            }
        }
        
        let mut buf = Vec::new();
        doc.build().pack(&mut std::io::Cursor::new(&mut buf))?;
        Ok(buf)
    }

    fn replace_placeholders(&self, text: &str, bill: &MerchantBill) -> String {
        let datetime = Local::now();
        let mut result = text.to_string();
        
        // 替换商家信息
        result = result.replace("{merchant_name}", &bill.merchant_name);
        result = result.replace("{year}", &datetime.year().to_string());
        result = result.replace("{month}", &datetime.month().to_string());
        
        // 替换水表读数
        result = result.replace("{prev_water_reading}", &bill.prev_water_reading.to_string());
        result = result.replace("{curr_water_reading}", &bill.curr_water_reading.to_string());
        
        // 替换用量计算
        result = result.replace("{water_usage}", &bill.water_usage.to_string());
        result = result.replace("{electricity_usage}", &bill.electricity_usage.to_string());
        
        // 替换费用计算
        result = result.replace("{water_unit_price}", &format!("{:.2}", bill.water_unit_price));
        result = result.replace("{electricity_unit_price}", &format!("{:.2}", bill.electricity_unit_price));
        result = result.replace("{water_amount}", &format!("{:.2}", bill.water_amount));
        result = result.replace("{electricity_amount}", &format!("{:.2}", bill.electricity_amount));
        result = result.replace("{total_amount}", &format!("{:.2}", bill.total_fee));
        
        // 替换电表详细信息
        if result.contains("{electricity_details}") {
            let details = bill.get_electricity_details().join("\n");
            result = result.replace("{electricity_details}", &details);
        }
        
        // 替换电表数量
        result = result.replace("{electricity_meter_count}", &bill.electricity_meters.len().to_string());
        
        result
    }
}
