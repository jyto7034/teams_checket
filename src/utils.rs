// src/utils.rs

use std::{
    collections::HashMap,
    error::Error,
    fs::{self, File},
    io::{self, BufRead, BufReader, BufWriter, Write},
    path::{Path, PathBuf},
    sync::Once,
};
use tracing::{debug, error, info, warn};
use tracing_appender::rolling::{RollingFileAppender, Rotation};
use tracing_subscriber::{EnvFilter, fmt, layer::SubscriberExt, util::SubscriberInitExt};

use crate::consts::DATE_FORMAT;

#[derive(Debug)]
pub struct Config {
    pub excel_path: PathBuf,
    pub manage_games: Vec<String>,
    pub notification_title: Option<String>,
    pub notification_message_template: Option<String>,
}

// 실행 파일 위치 가져오기
pub fn get_executable_dir() -> Result<PathBuf, Box<dyn Error>> {
    let exe_path = std::env::current_exe()?;
    Ok(exe_path
        .parent()
        .ok_or("실행 파일 경로를 찾을 수 없습니다.")?
        .to_path_buf())
}

// config.cfg 파일 읽기
pub fn read_config(path: &Path) -> Result<Config, Box<dyn Error>> {
    if !path.exists() {
        error!("설정 파일({})을 찾을 수 없습니다.", path.display());
        return Err(format!("설정 파일({})을 찾을 수 없습니다.", path.display()).into());
    }

    let file = File::open(path)?;
    let reader = BufReader::new(file);

    let mut excel_path_str = None;
    let mut manage_games = Vec::new();
    let mut notification_title = None;
    let mut notification_message_template = None;
    let mut current_section = "".to_string();

    for line in reader.lines() {
        let line = line?.trim().to_string();
        if line.is_empty() || line.starts_with(';') || line.starts_with('#') {
            continue;
        }

        if line.starts_with('[') && line.ends_with(']') {
            current_section = line[1..line.len() - 1].trim().to_lowercase();
            continue;
        }

        match current_section.as_str() {
            "target_path" => {
                if excel_path_str.is_none() {
                    excel_path_str = Some(line);
                } else {
                    warn!(
                        "[target_path]에 여러 경로가 지정됨. 첫 번째 경로만 사용: {}",
                        excel_path_str.as_ref().unwrap()
                    );
                }
            }
            "manage_game" => {
                manage_games.push(line);
            }
            "title" => {
                if notification_title.is_none() {
                    notification_title = Some(line);
                } else {
                    warn!("[title]에 여러 줄이 지정됨. 첫 번째 줄만 사용합니다.");
                }
            }
            "message" => {
                if notification_message_template.is_none() {
                    notification_message_template = Some(line);
                } else {
                    warn!("[message]에 여러 줄이 지정됨. 첫 번째 줄만 사용합니다.");
                }
            }
            _ => {} // 다른 섹션 무시
        }
    }

    let excel_path = match excel_path_str {
        Some(p) => PathBuf::from(p),
        None => {
            error!("설정 파일에 [target_path] 섹션 또는 경로가 없습니다.");
            return Err("설정 파일에 [target_path] 섹션 또는 경로가 없습니다.".into());
        }
    };

    if manage_games.is_empty() {
        warn!("설정 파일에 [manage_game] 섹션 또는 관리할 게임 이름이 없습니다.");
        return Err("설정 파일에 [manage_game] 섹션 또는 관리할 게임 이름이 없습니다.".into());
    }

    Ok(Config {
        excel_path,
        manage_games,
        notification_title,
        notification_message_template,
    })
}

pub fn excel_date_to_string(serial_date: f64) -> String {
    use chrono::{Duration, NaiveDate};
    let excel_epoch = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let days = serial_date.trunc() as i64;
    let date = excel_epoch + Duration::days(days);
    date.format(DATE_FORMAT).to_string()
}

pub fn excel_time_to_string(serial_time: f64) -> String {
    let total_seconds = (serial_time * 86400.0).round() as u32;
    let hours = total_seconds / 3600;
    let minutes = (total_seconds % 3600) / 60;
    let seconds = total_seconds % 60;
    format!("{:02}:{:02}:{:02}", hours, minutes, seconds)
}

pub fn write_missing_report(
    path: &Path,
    missing_data: &HashMap<String, Vec<String>>,
) -> Result<(), Box<dyn Error>> {
    // 기존 파일 삭제 시도
    if path.exists() {
        match fs::remove_file(path) {
            Ok(_) => info!("기존 보고서 파일 삭제: {:?}", path),
            Err(e) => {
                warn!("기존 보고서 파일 삭제 실패: {:?}, 오류: {}", path, e);
            }
        }
    }

    let file = File::create(path)?;
    let mut writer = BufWriter::new(file);

    if missing_data.is_empty() {
        info!("보고서 파일 작성: 누락된 항목이 없습니다.");
        writeln!(writer, "누락된 알림 처리 항목이 없습니다.")?;
    } else {
        info!("누락된 항목 보고서 작성 시작...");
        let mut sorted_sheets: Vec<&String> = missing_data.keys().collect();
        sorted_sheets.sort();

        for sheet_name in sorted_sheets {
            if let Some(entries) = missing_data.get(sheet_name) {
                writeln!(writer, "[{}]", sheet_name)?;
                debug!(
                    "시트 '{}'의 누락 항목 {}개 작성 중...",
                    sheet_name,
                    entries.len()
                );
                for entry in entries {
                    writeln!(writer, "{}", entry)?;
                }
                writeln!(writer)?;
            }
        }
        info!("누락된 항목 보고서 작성 완료: {:?}", path);
    }

    writer.flush()?;
    Ok(())
}

static INIT: Once = Once::new();
static mut GUARD: Option<tracing_appender::non_blocking::WorkerGuard> = None;
pub fn setup_logger() {
    INIT.call_once(|| {
        // 1. 파일 로거 설정
        let file_appender = RollingFileAppender::new(Rotation::DAILY, "logs", "app.log");
        let (non_blocking_file_writer, _guard) = tracing_appender::non_blocking(file_appender);

        // 2. 로그 레벨 필터 설정 (환경 변수 또는 기본값 INFO)
        let filter = EnvFilter::try_from_default_env().unwrap_or_else(|_| EnvFilter::new("info")); // 기본 INFO 레벨

        // 3. 콘솔 출력 레이어 설정
        let console_layer = fmt::layer()
            .with_writer(io::stdout) // 표준 출력으로 설정
            .with_ansi(true) // ANSI 색상 코드 사용 (터미널 지원 시)
            .with_thread_ids(true) // 스레드 ID 포함
            .with_thread_names(true) // 스레드 이름 포함
            .with_file(true) // 파일 경로 포함
            .with_line_number(true) // 라인 번호 포함
            .with_target(false) // target 정보 제외 (선택 사항)
            .pretty(); // 사람이 읽기 좋은 포맷

        // 4. 파일 출력 레이어 설정
        let file_layer = fmt::layer()
            .with_writer(non_blocking_file_writer) // Non-blocking 파일 로거 사용
            .with_ansi(false) // 파일에는 ANSI 코드 제외
            .with_thread_ids(true)
            .with_thread_names(true)
            .with_file(true)
            .with_line_number(true)
            .with_target(false);

        // 5. 레지스트리(Registry)에 필터와 레이어 결합
        tracing_subscriber::registry()
            .with(filter) // 필터를 먼저 적용
            .with(console_layer) // 콘솔 레이어 추가
            .with(file_layer) // 파일 레이어 추가
            .init(); // 전역 Subscriber로 설정

        unsafe {
            GUARD = Some(_guard);
        }

        tracing::info!("로거 초기화 완료: 콘솔 및 파일(logs/app.log) 출력 활성화.");
    });
}
