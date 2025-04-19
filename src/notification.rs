// src/notification.rs
use std::{collections::HashMap, error::Error, path::PathBuf, process::Command};

use calamine::{DataType, Reader, Xlsx, open_workbook};
// --- chrono::NaiveTime 추가 ---
use chrono::{Local, NaiveTime, Timelike};
// --- Duration도 chrono에서 직접 사용 ---
use chrono::Duration as ChronoDuration;
use tokio::time::{Duration, sleep};
use tracing::{debug, error, info, warn};

use crate::{
    consts::{DATE_FORMAT, OUTPUT_FILE_NAME},
    utils::{Config, excel_date_to_string, excel_time_to_string, write_missing_report},
};

pub type NotificationList = HashMap<String, Vec<String>>;

fn check_for_missed_notifications(config: &Config) -> Result<NotificationList, Box<dyn Error>> {
    info!("누락 알림 확인 시작 (오늘 날짜 & 과거 시간 & 9분 경과 미완료 항목 확인)");
    let mut excel: Xlsx<_> = open_workbook(&config.excel_path).map_err(|e| {
        error!("엑셀 파일 열기 실패: {}", e);
        e
    })?;

    let now = Local::now();
    let today_str = now.format(DATE_FORMAT).to_string();
    let current_naive_time = now.time();

    info!(
        "오늘 날짜: {}, 현재 시각: {}",
        today_str,
        current_naive_time.format("%H:%M:%S")
    );

    let mut missing_notifications: NotificationList = HashMap::new();
    let grace_period = ChronoDuration::minutes(9);

    for sheet_name in &config.manage_games {
        debug!(" - 시트 '{}' 확인 중...", sheet_name);
        match excel.worksheet_range(sheet_name) {
            Ok(range) => {
                let mut current_sheet_missing = Vec::new();
                let mut row_num = 0;

                for row in range.rows() {
                    row_num += 1;

                    // B열: 날짜 추출
                    let date_cell = row.get(1);
                    let date_str = match date_cell {
                        Some(DataType::String(s)) => Some(s.trim().to_string()),
                        Some(DataType::Float(f)) => Some(excel_date_to_string(*f)),
                        Some(DataType::DateTime(dt)) => Some(excel_date_to_string(*dt)),
                        Some(other_type) if !other_type.is_empty() => {
                            warn!(
                                "시트 '{}' 행 {} B열 예상 외 타입: {:?}, 처리 시도 중...",
                                sheet_name, row_num, other_type
                            );
                            if let Some(f_val) = other_type.as_f64() {
                                Some(excel_date_to_string(f_val))
                            } else {
                                warn!(
                                    "시트 '{}' 행 {} B열 {:?} 타입은 날짜로 처리 불가",
                                    sheet_name, row_num, other_type
                                );
                                None
                            }
                        }
                        _ => None,
                    };

                    // D열: 완료 여부 확인
                    let completed_cell = row.get(3);
                    let is_completed = match completed_cell {
                        Some(DataType::Empty) => false,
                        Some(DataType::String(s)) if s.trim().is_empty() => false,
                        Some(_) => true,
                        None => false,
                    };

                    // --- 조건 1 & 2: 오늘 날짜이고, 완료되지 않았는가? ---
                    if let Some(date) = date_str {
                        if date == today_str && !is_completed {
                            // --- 조건 3 & 4 를 위한 시간 처리 ---
                            let time_cell = row.get(2);
                            let time_str_opt = match time_cell {
                                // Option<String>으로 받기
                                Some(DataType::String(s)) => Some(s.trim().to_string()),
                                Some(DataType::Float(f)) => Some(excel_time_to_string(*f)),
                                Some(DataType::DateTime(dt)) => Some(excel_time_to_string(*dt)),
                                Some(other_type) if !other_type.is_empty() => {
                                    warn!(
                                        "시트 '{}' 행 {} C열 예상 외 타입: {:?}, 처리 시도 중...",
                                        sheet_name, row_num, other_type
                                    );
                                    if let Some(f_val) = other_type.as_f64() {
                                        Some(excel_time_to_string(f_val))
                                    } else {
                                        warn!(
                                            "시트 '{}' 행 {} C열 {:?} 타입은 시간으로 처리 불가",
                                            sheet_name, row_num, other_type
                                        );
                                        None
                                    }
                                }
                                _ => None,
                            };

                            if let Some(time_str) = time_str_opt {
                                // C열 시간 문자열을 NaiveTime으로 파싱 시도
                                match NaiveTime::parse_from_str(&time_str, "%H:%M:%S") {
                                    Ok(row_naive_time) => {
                                        // --- 조건 3: 과거 시간인가? ---
                                        if row_naive_time < current_naive_time {
                                            // --- 조건 4: 10분 유예 기간이 지났는가? ---
                                            let time_difference =
                                                current_naive_time - row_naive_time;
                                            if time_difference >= grace_period {
                                                // 모든 조건 충족! 누락 항목으로 추가
                                                let missing_entry =
                                                    format!("{} {}", date, time_str);
                                                debug!(
                                                    "  -> 누락 발견 (조건 충족): {}",
                                                    missing_entry
                                                );
                                                current_sheet_missing.push(missing_entry);
                                            } else {
                                                // 10분 유예 기간 중, 아직 누락 아님
                                                debug!(
                                                    "  -> 누락 건너뜀 (10분 유예 기간): {} {}",
                                                    date, time_str
                                                );
                                            }
                                        } else {
                                            // 미래 시간이므로 대상 아님
                                        }
                                    }
                                    Err(e) => {
                                        // 시간 파싱 실패 시 경고 로그
                                        warn!(
                                            "행 {} C열 시간 형식 파싱 오류 '{}': {}",
                                            row_num, time_str, e
                                        );
                                    }
                                }
                            } else {
                                // C열에 시간 정보 자체가 없는 경우 경고
                                warn!("행 {} C열에 시간 정보 없음. 누락 검사에서 제외.", row_num);
                            }
                        } // if date == today_str && !is_completed
                    } // if let Some(date) = date_str
                } // 행 반복 종료

                if !current_sheet_missing.is_empty() {
                    missing_notifications.insert(sheet_name.clone(), current_sheet_missing);
                }
            } // Ok(range)
            Err(e) => {
                error!("시트 '{}' 범위 읽기 오류: {}", sheet_name, e);
            }
        } // match result
    } // 시트 반복 종료

    // 결과 로그 메시지
    if missing_notifications.is_empty() {
        info!("확인 결과: 조건에 맞는 누락된 알림 항목 없음");
    } else {
        let total_missing_count: usize = missing_notifications.values().map(|v| v.len()).sum();
        info!(
            "확인 결과: 총 {}개의 누락된 알림 항목 발견 ({}개 시트)",
            total_missing_count,
            missing_notifications.len()
        );
    }

    Ok(missing_notifications)
}

pub async fn start_notification_service(
    config: &Config,
    exe_dir: &PathBuf,
) -> Result<(), Box<dyn Error>> {
    info!("알림 확인 서비스 시작. 매시간 11, 26, 41, 56분에 실행됩니다.");
    let output_path = exe_dir.join(OUTPUT_FILE_NAME);
    let notification_exe_path = exe_dir.join("notification.exe");

    loop {
        let now = Local::now();
        let current_minute = now.minute();

        let trigger_check = match current_minute {
            11 | 26 | 41 | 56 => true,
            _ => false,
        };

        if trigger_check {
            info!(
                "현재 시간: {}, 실행 조건 충족. 누락 항목 검사 시작...",
                now.format("%H:%M:%S")
            );
            match check_for_missed_notifications(config) {
                Ok(notification_list) => {
                    if !notification_list.is_empty() {
                        let total_missing_count: usize =
                            notification_list.values().map(|v| v.len()).sum();
                        info!(
                            "{}개 시트에서 총 {}개의 누락된 항목 발견.",
                            notification_list.len(),
                            total_missing_count
                        );
                        for (sheet, entries) in &notification_list {
                            let entries_str = entries.join(", ");
                            info!("  - 시트 [{}]: {}", sheet, entries_str);
                        }

                        if let Err(e) = write_missing_report(&output_path, &notification_list) {
                            error!("missing.txt 파일 쓰기 실패: {}", e);
                        } else {
                            info!("누락 목록을 {} 에 저장했습니다.", output_path.display());
                        }

                        let title = config.notification_title.as_deref().unwrap_or("알림");

                        let message_template = config
                            .notification_message_template
                            .as_deref()
                            .unwrap_or("{count}개의 누락된 데이터가 존재합니다!");
                        let message =
                            message_template.replace("{count}", &total_missing_count.to_string());

                        info!("알림 실행: Title='{}', Message='{}'", title, message);

                        if notification_exe_path.exists() {
                            match Command::new(notification_exe_path.clone())
                                .arg("--title")
                                .arg(title)
                                .arg("--message")
                                .arg(&message)
                                .status()
                            {
                                Ok(status) => {
                                    if status.success() {
                                        info!("notification.exe 실행 성공.");
                                    } else {
                                        warn!(
                                            "notification.exe 실행 완료되었으나, 성공 상태가 아님: {:?}",
                                            status.code()
                                        );
                                    }
                                }
                                Err(e) => {
                                    error!("notification.exe 실행 실패: {}", e);
                                }
                            }
                        } else {
                            warn!(
                                "notification.exe 파일을 찾을 수 없습니다: {}",
                                notification_exe_path.display()
                            );
                        }
                    }
                }
                Err(e) => {
                    error!("알림 확인 중 오류 발생: {}", e);
                    sleep(Duration::from_secs(60)).await;
                    continue;
                }
            }
            info!("다음 확인 시간까지 대기합니다 (약 65초 후 재검사)...");
            sleep(Duration::from_secs(65)).await;
        } else {
            let seconds_until_next_minute = 60 - now.second();
            let sleep_duration_secs = if current_minute == 10
                || current_minute == 25
                || current_minute == 40
                || current_minute == 55
            {
                1
            } else {
                (seconds_until_next_minute % 60).max(1)
            };
            sleep(Duration::from_secs(sleep_duration_secs as u64)).await;
        }
    }
    // TODO: 기능 수행 대기 시간동안, command 를 입력받을 수 있게 해야함.
    // 1. quit: 종료
    // 2. check: 즉시 확인
    // 3. status: 현재 상태 확인
    // 4. help: 도움말 출력
    // 5. log: 로그 출력 (로그 레벨 조정 필요)
    // 근데 따로 job spawn 하고.. command parsing 받아야 하고.. 귀찬다
    // Ok(())
}
