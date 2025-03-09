import logging

log_file = "log_output.log"

# ログ設定（ファイルとコンソールの両方に出力）
logging.basicConfig(
    level=logging.INFO, 
    format="[%(asctime)s] - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file, mode="a", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

def log_message(message: str):
    """メッセージをログファイルに記録し、同時にprintする"""
    logging.info("")
    logging.info("########################################")
    logging.info(f"# {message}")
    logging.info("########################################")

def log_less_message(message: str):
    """メッセージをログファイルに記録し、同時にprintする"""
    logging.info(f"{message}")

if __name__ == "__main__":
    # 使用例
    log_message("このメッセージはログに記録され、同時に表示されます。")
