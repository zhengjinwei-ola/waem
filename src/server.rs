use std::{io::Write, fs::File, path::PathBuf};
use axum::{response::{Html, IntoResponse}, routing::{get, post}, Router, extract::Multipart};
use anyhow::Result;
use tempfile::tempdir;

// 导入库crate（同包名）的导出项
use water_and_electricity_meter::{HeadersMap, read_data_file, generate_word_document_with_template, GenerateOptions};

#[tokio::main]
async fn main() -> Result<()> {
    let app = Router::new()
        .route("/", get(index))
        .route("/upload", post(upload));

    let port = std::env::var("PORT").unwrap_or_else(|_| "3002".to_string());
    let addr = format!("0.0.0.0:{}", port);
    
    println!("🚀 Excel到Word转换器服务启动中...");
    println!("📍 服务地址: http://{}", addr);
    println!("📝 上传Excel/CSV文件到: http://{}/", addr);
    
    let listener = tokio::net::TcpListener::bind(&addr).await.unwrap();
    println!("✅ 服务启动成功！");
    
    axum::serve(listener, app).await?;
    Ok(())
}

async fn index() -> impl IntoResponse {
    Html(r#"<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<title>水电表生成系统</title>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<style>
body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif;padding:24px;}
.card{max-width:680px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;padding:24px;box-shadow:0 10px 25px rgba(0,0,0,0.05)}
label{display:block;margin:12px 0 6px;color:#374151}
input[type=file],input[type=text]{width:100%;padding:10px;border:1px solid #d1d5db;border-radius:8px}
button{margin-top:16px;padding:10px 16px;background:#2563eb;color:white;border:none;border-radius:8px;cursor:pointer}
small{color:#6b7280}
</style>
</head>
<body>
<div class="card">
  <h2>水电表生成系统</h2>
  <form action="/upload" method="post" enctype="multipart/form-data">
    <label>选择文件（.xlsx 或 .csv）</label>
    <input name="file" type="file" accept=".xlsx,.csv" required />
    <!-- 店铺名称列从CSV获取，不在页面展示 -->
    <label>自定义标题（可选，默认：yyyy年MM月抄表计费通知单）</label>
    <input name="custom_title" type="text" placeholder="例如：2025年08月抄表计费通知单"/>
    <label>每页表格数量（默认 3）</label>
    <input name="per_page" type="text" value="3"/>
    <label>抄表人</label>
    <input name="meter_reader" type="text" placeholder="请输入抄表人"/>
    <label>抄表日期</label>
    <input name="meter_date" type="text" placeholder="例如：2025年08月16日"/>
    <button type="submit">生成Word</button>
    <div><small>提示：表头需要与输入框一致或为常见别名。</small></div>
  </form>
</div>
</body>
</html>"#)
}

async fn upload(mut multipart: Multipart) -> impl IntoResponse {
    let mut params = DefaultParams::default();
    let mut saved_path: Option<PathBuf> = None;

    while let Ok(Some(field)) = multipart.next_field().await {
        let name = field.name().map(|s| s.to_string()).unwrap_or_default();
        if name == "file" {
            let orig_name: String = field.file_name().map(|s| s.to_string()).unwrap_or_else(|| "upload".to_string());
            let bytes = field.bytes().await.unwrap_or_default();
            // preserve extension for type detection
            let dir = tempdir().unwrap();
            let ext = std::path::Path::new(&orig_name).extension().and_then(|e| e.to_str()).unwrap_or("");
            let fname = if ext.is_empty() { "upload.csv".to_string() } else { orig_name.clone() };
            let path = dir.path().join(fname);
            let mut f = File::create(&path).unwrap();
            f.write_all(&bytes).unwrap();
            saved_path = Some(path);
            // keep dir alive until function end by moving it into path parent? We'll leak dir by forgetting it to keep file.
            std::mem::forget(dir);
            println!("received file: {} ({} bytes)", orig_name, bytes.len());
        } else {
            let value = field.text().await.unwrap_or_default();
            match name.as_str() {
                "prev_e" => params.prev_e = value,
                "curr_e" => params.curr_e = value,
                "prev_w" => params.prev_w = value,
                "curr_w" => params.curr_w = value,
                "water_price" => params.water_price = value,
                "elec_price" => params.elec_price = value,
                "meter_reader" => params.meter_reader = value,
                "meter_date" => params.meter_date = value,
                "custom_title" => params.custom_title = value,
                "per_page" => params.per_page = value,
                _ => {}
            }
        }
    }

    let path = if let Some(p) = saved_path { p } else { return Html("上传失败：未收到文件").into_response() };

    match process_file_to_docx(path, params).await {
        Ok((filename, bytes)) => (
            [("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
             ("Content-Disposition", &format!("attachment; filename=\"{}\"", filename))],
            bytes
        ).into_response(),
        Err(e) => Html(format!("生成失败：{}", e)).into_response(),
    }
}

#[derive(Default)]
struct DefaultParams {
    prev_e: String,
    curr_e: String,
    prev_w: String,
    curr_w: String,
    water_price: String,
    elec_price: String,
    meter_reader: String,
    meter_date: String,
    custom_title: String,
    per_page: String,
}

async fn process_file_to_docx(path: PathBuf, params: DefaultParams) -> anyhow::Result<(String, Vec<u8>)> {
    use anyhow::Context;
    
    // 创建新的HeadersMap结构
    let headers = HeadersMap {
        merchant: "店铺名称",
        prev_e: &params.prev_e,
        curr_e: &params.curr_e,
        prev_w: &params.prev_w,
        curr_w: &params.curr_w,
        w_price: &params.water_price,
        e_price: &params.elec_price,
        electricity_price: &params.elec_price,
        electricity_prefix: "电表",
        water_electricity_labor_fee: "水电人工费",
        garbage_disposal_fee: "垃圾处理费",
    };

    // 直接调用main.rs中的函数
    let mut bills = read_data_file(path.to_str().unwrap(), &headers)
        .with_context(|| "解析数据失败")?;
    if bills.is_empty() { anyhow::bail!("文件中没有有效数据"); }

    // 将抄表人和抄表日期写入每条记录
    for bill in bills.iter_mut() {
        bill.set_meter_info(
            if params.meter_reader.trim().is_empty() { None } else { Some(params.meter_reader.clone()) },
            if params.meter_date.trim().is_empty() { None } else { Some(params.meter_date.clone()) },
        );
    }

    // 生成Word文档
    let per_page = params.per_page.trim().parse::<usize>().unwrap_or(1);
    let opts = GenerateOptions { custom_title: if params.custom_title.trim().is_empty() { None } else { Some(params.custom_title.clone()) }, per_page };
    let docx_content = generate_word_document_with_template(&bills, Some(opts))
        .map_err(|e| anyhow::anyhow!("生成Word文档失败: {}", e))?;

    let now = chrono::Local::now();
    let filename = if params.custom_title.trim().is_empty() {
        format!("report_{}{}.docx", now.format("%m"), now.format("%Y"))
    } else {
        // 使用自定义标题作为文件名，移除特殊字符
        let clean_title = params.custom_title
            .replace("年", "")
            .replace("月", "")
            .replace("日", "")
            .replace(" ", "_")
            .replace("/", "_")
            .replace("\\", "_")
            .replace(":", "_")
            .replace("*", "_")
            .replace("?", "_")
            .replace("\"", "_")
            .replace("<", "_")
            .replace(">", "_")
            .replace("|", "_");
        format!("{}.docx", clean_title)
    };
    Ok((filename, docx_content))
}

