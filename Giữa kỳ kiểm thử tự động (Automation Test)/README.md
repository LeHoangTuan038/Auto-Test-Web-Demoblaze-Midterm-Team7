# Kiểm thử tự động Demoblaze

Bộ kiểm thử này dùng **Selenium WebDriver + Python + pytest** cho website [Demoblaze](https://demoblaze.com/). Đây là bộ Selenium chạy riêng, không dùng Playwright. Selenium Manager tự xử lý ChromeDriver; máy chạy test cần cài Google Chrome.

## Cài đặt

Mở terminal tại thư mục này:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Chạy test và tạo báo cáo

```powershell
pytest
pytest --headed
pytest -m smoke
pytest -m regression
pytest --html=results/pytest-html-report.html --self-contained-html
```

Mặc định Chrome chạy ẩn (headless). Dùng `--headed` khi cần trình diễn trực tiếp. Mỗi lần pytest kết thúc, `results/report.html` được tạo dưới dạng HTML tĩnh để bảng kết quả vẫn hiển thị khi trình duyệt chặn JavaScript. Báo cáo tương tác của pytest-html (nếu dùng tùy chọn `--html`) lưu riêng tại `results/pytest-html-report.html`.

## Cấu trúc

```text
.
├── tests/
│   ├── helpers.py
│   ├── test_login.py
│   ├── test_cart.py
│   └── test_checkout.py
├── conftest.py
├── pytest.ini
├── requirements.txt
├── README.md
├── Test-Cases-Demoblaze.xlsx
└── Midterm-Automation-Test.pptx
```

`conftest.py` khởi tạo và đóng Chrome cho từng test. `tests/helpers.py` chứa các thao tác dùng chung như đăng ký, đăng nhập, thêm sản phẩm và mở checkout. Các tài khoản kiểm thử được tạo với username duy nhất mỗi lần chạy.

`Test-Cases-Demoblaze.xlsx` là bảng test case 16 dòng; `Midterm-Automation-Test.pptx` là slide thuyết trình 15 trang. Có thể tạo lại hai tệp này sau khi chỉnh sửa dữ liệu trong `scripts/create_deliverables.py` bằng lệnh:

```powershell
python scripts/create_deliverables.py
```

## Danh sách 16 kịch bản kiểm thử

| ID | Loại | Kịch bản | Hàm test |
|---|---|---|---|
| TC01 | Smoke | Đăng ký và đăng nhập thành công | `test_sign_up_and_login_successfully` |
| TC02 | Smoke | Hiển thị danh sách sản phẩm | `test_product_catalog_is_displayed` |
| TC03 | Smoke | Thêm một sản phẩm vào giỏ | `test_user_can_add_a_product_to_cart` |
| TC04 | Smoke | Checkout thành công | `test_user_can_complete_purchase` |
| TC05 | Smoke | Đăng xuất | `test_user_can_log_out` |
| TC06 | Regression | Đăng nhập với sai mật khẩu | `test_login_rejects_wrong_password` |
| TC07 | Regression | Bỏ trống username | `test_login_requires_username` |
| TC08 | Regression | Bỏ trống password | `test_login_requires_password` |
| TC09 | Regression | Đăng ký username đã tồn tại | `test_sign_up_rejects_duplicate_username` |
| TC10 | Regression | Lọc danh mục Phones | `test_category_filter_shows_phones` |
| TC11 | Regression | Chuyển trang danh sách sản phẩm | `test_user_can_navigate_to_next_catalog_page` |
| TC12 | Regression | Xóa sản phẩm cuối cùng khỏi giỏ | `test_removing_last_product_leaves_cart_empty` |
| TC13 | Regression | Checkout bỏ trống Name | `test_purchase_requires_name` |
| TC14 | Regression | Checkout bỏ trống Credit card | `test_purchase_requires_credit_card` |
| TC15 | Regression | Giỏ trống và mở form đặt hàng | `test_empty_cart_has_no_products_and_order_form_can_be_opened` |
| TC16 | Regression | Tổng giỏ hàng bằng tổng giá sản phẩm | `test_cart_total_matches_sum_of_item_prices` |

### Các điểm khác với SauceDemo

Tài liệu giữa kỳ có một số kịch bản chỉ phù hợp với SauceDemo. Demoblaze không có tài khoản khóa, tùy chọn sắp xếp, ZIP code, badge giỏ hàng hoặc dòng thuế. Bộ test thay các case đó bằng kiểm tra username đăng ký trùng, lọc danh mục và chuyển trang, validation Credit card, trạng thái giỏ trống và tổng giá sản phẩm. Test TC15 chỉ kiểm tra trạng thái giỏ trống và khả năng mở form; không gửi đơn hàng trống vì trang vẫn có thể báo mua thành công.
