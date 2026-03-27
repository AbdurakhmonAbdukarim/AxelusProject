import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from modules.telegram_bot import run_bot
    print("Starting Telegram bot...")
    run_bot()


if __name__ == "__main__":
    main()