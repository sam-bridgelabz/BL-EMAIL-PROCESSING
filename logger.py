import sys
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
    datefmt="%Y-%m-%d %H:%M",
)

logger = logging.getLogger(__name__)
