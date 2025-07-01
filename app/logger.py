import sys
from datetime import datetime
import os

from loguru import logger as _logger

from app.config import PROJECT_ROOT


_print_level = "INFO"


def define_log_level(print_level="INFO", logfile_level="DEBUG", name: str = None):
    """Adjust the log level to above level"""
    global _print_level
    _print_level = print_level

    current_date = datetime.now()
    formatted_date = current_date.strftime("%Y%m%d%H%M%S")
    log_name = (
        f"{name}_{formatted_date}" if name else formatted_date
    )  # name a log with prefix name

    _logger.remove()
    _logger.add(sys.stderr, level=print_level)
    _logger.add(PROJECT_ROOT / f"logs/{log_name}.log", level=logfile_level)
    return _logger


logger = define_log_level()


def write_trace_md(content: str, filename: str = None):
    """
    Write detailed trace info to a markdown file in logs/trace/.
    If filename is None, use trace_时间戳.md
    """
    trace_dir = PROJECT_ROOT / "logs/trace"
    os.makedirs(trace_dir, exist_ok=True)
    if filename is None:
        filename = f"trace_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
    file_path = trace_dir / filename
    with open(file_path, "a", encoding="utf-8") as f:
        f.write(content + "\n")


if __name__ == "__main__":
    logger.info("Starting application")
    logger.debug("Debug message")
    logger.warning("Warning message")
    logger.error("Error message")
    logger.critical("Critical message")

    try:
        raise ValueError("Test error")
    except Exception as e:
        logger.exception(f"An error occurred: {e}")
