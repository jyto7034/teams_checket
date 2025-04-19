// src/main.rs
use log::{error, info}; // 또는 tracing 사용
use std::error::Error;

use checker::{
    consts::CONFIG_FILE_NAME,
    notification::start_notification_service,
    utils::{get_executable_dir, read_config, setup_logger},
    // validation 모듈 임포트는 이제 notification 모듈에서 사용
};

#[tokio::main]
async fn main() -> Result<(), Box<dyn Error>> {
    setup_logger();

    info!("팀즈 알림 누락 주기적 검사 도구를 시작합니다...");

    let exe_dir = get_executable_dir()?; // exe_dir 얻기
    info!("실행 파일 디렉토리: {:?}", exe_dir);
    let config_path = exe_dir.join(CONFIG_FILE_NAME);

    info!("설정 파일 읽는 중: {:?}", config_path);
    let config = read_config(&config_path).map_err(|e| {
        error!("설정 파일 처리 중 오류 발생: {}", e);
        e
    })?;
    info!(" - Excel 경로: {}", config.excel_path.display());
    info!(" - 관리 대상 시트: {:?}", config.manage_games);
    info!(" - 알림 제목 (설정됨): {:?}", config.notification_title);
    info!(
        " - 알림 메시지 템플릿 (설정됨): {:?}",
        config.notification_message_template
    );

    if !config.excel_path.exists() {
        error!(
            "설정된 Excel 파일을 찾을 수 없습니다: {}",
            config.excel_path.display()
        );
        return Err(format!("Excel file not found: {}", config.excel_path.display()).into());
    }

    info!("주기적 알림 확인 서비스 시작...");
    if let Err(e) = start_notification_service(&config, &exe_dir).await {
        error!("알림 서비스 실행 중 심각한 오류 발생: {}", e);
        return Err(e);
    }

    info!("알림 확인 서비스가 정상적으로 종료되었습니다 (예상치 못한 상황).");
    Ok(())
}
