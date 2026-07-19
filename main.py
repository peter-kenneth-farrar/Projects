"""
main.py
Entry point — scrapes Microsoft Excel issues page and updates OneNote.
"""

import logging
import sys
from scraper import scrape
from onenote_updater import update_onenote

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)


def main():
    logger.info("=" * 60)
    logger.info("Microsoft Excel Issues → OneNote Updater")
    logger.info("=" * 60)

    # Step 1: Scrape
    logger.info("\n[Step 1/2] Scraping Microsoft support page...")
    try:
        data = scrape()
        logger.info(f"  Scraped {data['total_issues']} issues across {len(data['sections'])} sections")
    except Exception as e:
        logger.error(f"Scraping failed: {e}")
        sys.exit(1)

    # Step 2: Update OneNote
    logger.info("\n[Step 2/2] Updating OneNote page...")
    try:
        update_onenote(data)
        logger.info("  OneNote page updated successfully!")
    except Exception as e:
        logger.error(f"OneNote update failed: {e}")
        sys.exit(1)

    logger.info("\n" + "=" * 60)
    logger.info("✅ All done!")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()