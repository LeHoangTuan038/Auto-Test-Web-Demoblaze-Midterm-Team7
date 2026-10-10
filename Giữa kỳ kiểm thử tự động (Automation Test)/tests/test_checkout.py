import re

import pytest
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from helpers import (
    accept_alert,
    add_product_by_index,
    cart_rows,
    fill_order_form,
    open_cart,
    open_home,
    open_order_modal,
    place_order,
    wait_for,
)


@pytest.mark.smoke
def test_user_can_complete_purchase(browser):
    assert add_product_by_index(browser) == "Product added"
    open_cart(browser, expect_items=True)
    open_order_modal(browser)
    fill_order_form(browser)
    place_order(browser)

    confirmation = wait_for(
        browser, EC.visibility_of_element_located((By.CSS_SELECTOR, ".sweet-alert"))
    )
    assert "Thank you for your purchase!" in confirmation.text


@pytest.mark.regression
def test_purchase_requires_name(browser):
    assert add_product_by_index(browser) == "Product added"
    open_cart(browser, expect_items=True)
    open_order_modal(browser)
    fill_order_form(browser, name="")
    place_order(browser)

    assert accept_alert(browser) == "Please fill out Name and Creditcard."


@pytest.mark.regression
def test_purchase_requires_credit_card(browser):
    assert add_product_by_index(browser) == "Product added"
    open_cart(browser, expect_items=True)
    open_order_modal(browser)
    fill_order_form(browser, card="")
    place_order(browser)

    assert accept_alert(browser) == "Please fill out Name and Creditcard."


@pytest.mark.regression
def test_empty_cart_has_no_products_and_order_form_can_be_opened(browser):
    open_home(browser)
    open_cart(browser)

    assert not cart_rows(browser)
    assert browser.find_element(By.ID, "totalp").text.strip() in ("", "0")

    open_order_modal(browser)

    assert browser.find_element(By.ID, "orderModal").is_displayed()


@pytest.mark.regression
def test_cart_total_matches_sum_of_item_prices(browser):
    assert add_product_by_index(browser) == "Product added"
    open_cart(browser, expect_items=True)

    item_prices = []
    for row in cart_rows(browser):
        price_text = row.find_element(By.CSS_SELECTOR, "td:nth-child(3)").text
        price_match = re.search(r"\d+", price_text)
        assert price_match is not None
        item_prices.append(int(price_match.group()))
    cart_total = int(browser.find_element(By.ID, "totalp").text)

    assert item_prices
    assert cart_total == sum(item_prices)
