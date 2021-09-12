use std::env;
use std::path::Path;
use std::fs::File;
use std::io::{self, BufRead};
use chrono::prelude::*;
use serde::Deserialize;
use reqwest::Url;
use xlsxwriter::{FormatAlignment,FormatColor,FormatBorder,Workbook};

#[derive(Deserialize, Debug)]
struct InfoDetailData{
    uid: String,
    gacha_type: String,
    item_id: String,
    count: String,
    time: String,
    name: String,
    lang: String,
    item_type: String,
    rank_type: String,
    id: String,
}

#[derive(Deserialize, Debug)]
struct InfoData{
    page: String,
    size: String,
    total: String,
    list: Vec<InfoDetailData>,
    region: String,
}

#[derive(Deserialize, Debug)]
struct TypesInnerData{
    id: String,
    key: String,
    name: String,
}

#[derive(Deserialize, Debug)]
struct TypesData{
    gacha_type_list: Vec<TypesInnerData>,
}

#[derive(Deserialize, Debug)]
struct API<T>{
    retcode: i32,
    message: String,
    data: Option<T>,
}



fn read_lines<P>(filename: P) -> io::Result<io::Lines<io::BufReader<File>>>
where P: AsRef<Path>, {
    let file = File::open(filename)?;
    Ok(io::BufReader::new(file).lines())
}

fn get_url() -> Result<String,String> {
    let user_path = env::var("USERPROFILE").unwrap();
    let log_path = Path::new(&user_path).join(r#"AppData\LocalLow\miHoYo\原神\output_log.txt"#);

    let mut ret = String::new();
    if let Ok(lines) = read_lines(log_path) {
        for line in lines {
            if let Ok(log) = line {
                if log.starts_with("OnGetWebViewPageFinish") && log.ends_with("#/log") {
                    ret = log;
                }
            }
        }
    }
    if ret.is_empty() {
        return Err(String::from("没找到抽卡详情的请求，先在游戏里打开抽卡详情再导出数据"));
    }
    let mut url_offset = ret.find(':').unwrap()+1;
    ret.drain(..url_offset);
    url_offset = ret.find('?').unwrap();
    ret.replace_range(..url_offset, "https://hk4e-api.mihoyo.com/event/gacha_info/api/getGachaLog");
    Ok(ret)
}

fn get_info(url:&str) -> API<InfoData> {
    let resp = reqwest::blocking::get(url).unwrap()
    .json::<API<InfoData>>().unwrap();
    check_result(&resp).unwrap();
    resp
}

fn check_result<T: std::fmt::Debug>(data:&API<T>) -> Result<(), String>{
    match data.data{
        Some(_) => Ok(()),
        None => {
            Err(data.message.clone())
        }
    }
}

fn get_types(url:&str) -> Vec<TypesInnerData> {
    let url = url.replace("getGachaLog", "getConfigList");
    let resp = reqwest::blocking::get(&url).unwrap()
    .json::<API<TypesData>>().unwrap();
    check_result(&resp).unwrap();
    resp.data.unwrap().gacha_type_list

}

fn get_api(url:&str, size: i32, page: i32, gacha_type: String, end_id: String) -> API<InfoData> {
    let url = Url::parse_with_params(url, &[("size", size.to_string()), ("page", page.to_string()), ("gacha_type", gacha_type), ("lang", String::from("zh-cn")), ("end_id", end_id)]).unwrap();
    let resp = reqwest::blocking::get(url).unwrap()
    .json::<API<InfoData>>().unwrap();
    check_result(&resp).unwrap();
    resp
}

// 返回 页面列表<每页数据<数据>> 内外层都是逆序的
fn get_details(url: &str, gacha_type: &str) -> Vec<Vec<InfoDetailData>> {
    let mut i = 1;
    let mut end_id = "0".to_string();
    let mut ret: Vec<Vec<InfoDetailData>> = Vec::new();
    loop {
        print!("正在获取第{}页...", i);
        let records = get_api(url, 20, i, gacha_type.to_string(), end_id.clone());
        let details = records.data.unwrap();
        println!("OK");
        if details.list.is_empty() {
            break;
        }
        i += 1;
        let last = details.list.last().unwrap();
        end_id = last.id.clone();
        ret.push(details.list);
    }
    ret
}

fn set_content_format(workbook:&Workbook) -> xlsxwriter::Format{
    workbook 
        .add_format()
        .set_font_name("微软雅黑")
        .set_align(FormatAlignment::Left)
        .set_border_color(FormatColor::Custom(0xc4c2bf))
        .set_bg_color(FormatColor::Custom(0xebebeb))
        .set_border(FormatBorder::Thin)
}

fn write_xlsx(workbook:&Workbook, detail: &[Vec<InfoDetailData>], name: &str) {
    let mut sheet = workbook.add_worksheet(Some(name)).unwrap();
    let content_fmt = set_content_format(workbook);
    let title_fmt = workbook
        .add_format()
        .set_font_name("微软雅黑")
        .set_align(FormatAlignment::Left)
        .set_font_color(FormatColor::Custom(0x757575))
        .set_border_color(FormatColor::Custom(0xc4c2bf))
        .set_bg_color(FormatColor::Custom(0xdbd7d3))
        .set_bold()
        .set_border(FormatBorder::Thin);
    let start3_fmt = set_content_format(workbook)
        .set_font_color(FormatColor::Custom(0x8e8e8e));
    let start4_fmt = set_content_format(workbook)
        .set_bold()
        .set_font_color(FormatColor::Custom(0xa256e1));
    let start5_fmt = set_content_format(workbook)
        .set_bold()
        .set_font_color(FormatColor::Custom(0xbd6932));
    sheet.set_column(0, 0, 22.0, None).unwrap();
    sheet.set_column(1, 1, 14.0, None).unwrap();
    for (idx,title) in ["时间","名称","类别","星级","总次数","保底内"].iter().enumerate(){
        sheet.write_string(0, idx as u16, title, Some(&title_fmt)).unwrap();
    }
    sheet.freeze_panes(1, 0);
    let mut idx = 0;
    let mut pdx = 0;
    let mut i = 0;
    for page in detail.iter().rev(){
        for gacha in page.iter().rev() {
            idx += 1;
            pdx += 1;
            i += 1;
            let excel_data = [ &gacha.time, &gacha.name, &gacha.item_type];
            for (index, item) in excel_data.iter().enumerate(){
                sheet.write_string(i, index as u16, item, Some(&content_fmt)).unwrap();
            }
            sheet.write_number(i, 4, idx as f64, Some(&content_fmt)).unwrap();
            sheet.write_number(i, 5, pdx as f64, Some(&content_fmt)).unwrap();
            match gacha.rank_type.parse::<i32>().unwrap() {
                3 => sheet.write_number(i, 3, 3.0, Some(&start3_fmt)).unwrap(),
                4 => sheet.write_number(i, 3, 4.0, Some(&start4_fmt)).unwrap(),
                5 => {
                    sheet.write_number(i, 3, 5.0, Some(&start5_fmt)).unwrap();
                    pdx = 0;
                },
                _ => ()
            }
        }
    }
}

fn main() {
    let url = get_url();
    match url {
        Ok(url) =>{
            let info = get_info(&url);
            if let Err(e) = check_result(&info){
                println!("{}",e);
                return;
            }
            let types = get_types(&url);
            let now: DateTime<Local> = Local::now();
            let now = now.format("%Y%m%d%H%M%S");
            let xlsx_name = format!("gachaExport-{}.xlsx", now);
            let workbook = Workbook::new(&xlsx_name);
            for gacha_type in types {
                println!("正在获取 {} 的记录...", gacha_type.name);
                let res = get_details(&url, &gacha_type.key);
                write_xlsx(&workbook, &res, &gacha_type.name);
            }
            workbook.close().unwrap();
            println!("记录已写入 {} ", xlsx_name);
        },
        Err(e) => {
            println!("{}",e)
        }
    }


    
}
