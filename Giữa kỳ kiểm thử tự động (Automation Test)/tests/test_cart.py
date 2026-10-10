import pytest
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from helpers import (
    add_product_by_index,
    cart_rows,
    open_cart,
    open_home,
    wait_for,
)


@pytest.mark.smoke
def test_product_catalog_is_displayed(browser):
    open_home(browser)

    products = browser.find_elements(By.CSS_SELECTOR, "#tbodyid .card h4 a")
    assert products
    assert all(product.text.strip() for product in products)


@pytest.mark.smoke
def test_user_can_add_a_product_to_cart(browser):
    assert add_product_by_index(browser) == "Product added"
    open_cart(browser, expect_items=True)

    assert len(cart_rows(browser)) == 1


@pytest.mark.regression
def test_category_filter_shows_phones(browser):
    open_home(browser)
    initial_titles = [
        element.text.strip().lower()
        for element in browser.find_elements(By.CSS_SELECTOR, "#tbodyid .card h4 a")
    ]
    assert any("vaio" in title for title in initial_titles)

    browser.find_element(By.XPATH, "//a[normalize-space()='Phones']").click()
    wait_for(
        browser,
        lambda current_driver: current_driver.find_elements(
            By.CSS_SELECTOR, "#tbodyid .card h4 a"
        )
        and not any(
            "vaio" in element.text.lower()
            for element in current_driver.find_elements(
                By.CSS_SELECTOR, "#tbodyid .card h4 a"
            )
        ),
    )
    titles = [
        element.text.strip().lower()
        for element in browser.find_elements(By.CSS_SELECTOR, "#tbodyid .card h4 a")
    ]

    assert titles
    assert not any("vaio" in title for title in titles)


@pytest.mark.regression
def test_user_can_navigate_to_next_catalog_page(browser):
    open_home(browser)
    first_product = browser.find_element(
        By.CSS_SELECTOR, "#tbodyid .card h4 a"
    )
    first_page_title = first_product.text
    browser.find_element(By.ID, "next2").click()

    wait_for(browser, EC.staleness_of(first_product))
    wait_for(
        browser,
        lambda current_driver: current_driver.find_elements(
            By.CSS_SELECTOR, "#tbodyid .card h4 a"
        ),
    )
    next_page_title = browser.find_element(
        By.CSS_SELECTOR, "#tbodyid .card h4 a"
    ).text

    assert next_page_title != first_page_title


@pytest.mark.regression
def test_removing_last_product_leaves_cart_empty(browser):
    assert add_product_by_index(browser) == "Product added"
    open_cart(browser, expect_items=True)
    assert len(cart_rows(browser)) == 1

    browser.find_element(By.CSS_SELECTOR, "#tbodyid a[onclick^='deleteItem']").click()

    wait_for(browser, lambda current_driver: not cart_rows(current_driver))
    assert not cart_rows(browser)
    assert not browser.find_elements(By.CSS_SELECTOR, ".shopping_cart_badge")
