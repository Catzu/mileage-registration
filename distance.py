from playwright.sync_api import sync_playwright
import time

from_postcode = "Hogebankweg 1, 5331 RD"
to_postcode = input("Enter destination postcode: ")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    
    page.goto("https://www.anwb.nl/verkeer/routeplanner", timeout=60000)

    # Wait until the "Van" field is ready (based on label text fallback)
    page.wait_for_selector("input[placeholder='Bijv. Amsterdam, Dam']")

    # Selectors by placeholder text, since 'data-test-id' was not reliable
    from_input = page.locator("input[placeholder='Bijv. Amsterdam, Dam']").first
    from_input.fill(from_postcode)
    time.sleep(1)
    page.keyboard.press("ArrowDown")
    page.keyboard.press("Enter")

    # Tab to next field ("Naar")
    page.keyboard.press("Tab")
    page.keyboard.type(to_postcode)
    time.sleep(1)
    page.keyboard.press("ArrowDown")
    page.keyboard.press("Enter")

    # Click "Plan route"
    page.get_by_role("button", name="Plan route").click()

    # Wait for distance to appear
    page.wait_for_selector(".route-description", timeout=30000)

    # Get distance text
    distance_element = page.locator(".route-description").first
    distance_text = distance_element.inner_text()
    print(f"Distance from {from_postcode} to {to_postcode} (one-way): {distance_text}")

    browser.close()