import re
import time
import tkinter as tk
from tkinter import simpledialog
from playwright.sync_api import Playwright, sync_playwright, expect


 
def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("http://rb3qraa0.server.bosch.com:8001/sap/bc/gui/sap/its/webgui/#")

    print('Acessando LX15...')
    time.sleep(3)
    page.get_by_role("textbox", name="Enter transaction code").fill("LX15")
    page.get_by_role("textbox", name="Enter transaction code").press("Enter")
    page.locator("div").filter(has_text=re.compile(r"^Get Variant\.\.\.$")).click()
    time.sleep(3)
    page.get_by_role("row", name="MMS3CA", exact=True).locator("div").nth(1).click()
    page.get_by_title("Choose (F2)").click()


    ### utilizando atualmente depósito J0A como padrão
    # lógica pra acessar depósito:
    #     page.get_by_role("textbox", name="Storage Type").fill("J0A")
    #     page.get_by_role("textbox", name="Storage Type").press("Enter")
    ### nos testes, utilizar H10 ou J0A


    time.sleep(3)
    storage_code = "J0A"
    page.get_by_role("textbox", name="Storage Type").fill(storage_code)
    time.sleep(1)
    page.get_by_role("textbox", name="Storage Type").press("Enter")
    time.sleep(3)
    page.get_by_role("textbox", name="Enter transaction code").press("F8")


#SE DER A MSG DO PRINT, NEM DEIXA CONTINUAR. TRATAR ESSE ERRO
    print('AGUARDANDO 1min')
    page.screenshot(path="screenshot_storage.png")

    time.sleep(60)
    page.locator("div").filter(has_text=re.compile(r"^Activate$")).click()
    time.sleep(3)
    page.locator("div").filter(has_text=re.compile(r"^Exit$")).click()
    time.sleep(3)
    page.locator("div").filter(has_text=re.compile(r"^Exit$")).click()
    time.sleep(3)

    page.get_by_role("textbox", name="Enter transaction code").fill("SM35")
    page.get_by_role("textbox", name="Enter transaction code").press("Enter")
    time.sleep(3)
    page.locator(".urST5SCMetricInner").first.click()
    time.sleep(3)
    page.locator("div").filter(has_text=re.compile(r"^Process$")).click()
    time.sleep(3)
    page.get_by_text("Background", exact=True).click()
    time.sleep(3)
    page.locator("#SAPMSBDC_CC300_1-tbcontainer div").filter(has_text=re.compile(r"^Process$")).click()
    time.sleep(3)
    page.locator("div").filter(has_text=re.compile(r"^Exit$")).click()
    time.sleep(3)

    # ---------------------
    context.close()
    browser.close()
 
 
if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
