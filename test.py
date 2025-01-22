import os
import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 変数の初期化
v_LoginURL = "https://id.jobcan.jp/users/sign_in"
v_LogFilePath = "C:\\RPALogs\\JobcanUpdateLog.txt"
v_ContractList = []
v_Today = datetime.datetime.now()

# ログファイルの準備
def log_message(message):
    # ディレクトリが存在しない場合は作成
    os.makedirs(os.path.dirname(v_LogFilePath), exist_ok=True)
    with open(v_LogFilePath, "a") as log_file:
        log_file.write(f"[{datetime.datetime.now()}] {message}\n")

# ジョブカンにログイン
driver = webdriver.Chrome()  # ChromeDriverのパスを指定する場合は引数を追加
driver.get(v_LoginURL)

# ユーザーにIDとパスワードを入力してもらう
print("ジョブカンのログイン画面が開きました。IDとパスワードを入力してください。")
print("ログイン後、Enterキーを押してください。")

# ユーザーが手動でログインするのを待機
input("ログインが完了したら、Enterキーを押してください。")

# ログイン成功判定
time.sleep(3)  # ページが読み込まれるまで待機
if "ログイン" in driver.title:
    log_message("ログイン失敗")
    driver.quit()
    exit()

# ワークフロー画面へ移動
try:
    # WebDriverWaitを使用して「WF」リンクが表示されるまで待機
    wf_link = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "WF"))
    )
    wf_link.click()
    print("WFをクリックしました。")  # クリックしたことを出力

    # サイドバーの「自分の申請一覧」をXPathで直接クリック
    try:
        application_function_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[@go-click='/myrequests/']"))
        )
        application_function_link()
        print("自分の申請一覧をクリックしました。")  # クリックしたことを出力
    except Exception:
        print("自分の申請一覧が見つかりません。")  # 要素が見つからなかった場合のメッセージ
        log_message("自分の申請一覧が見つかりません。")
        driver.quit()
        exit()

    # 「すべて」をクリック
    all_link = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "すべて"))
    )
    all_link.click()
    print("すべてをクリックしました。")  # クリックしたことを出力

    # 契約一覧を取得
    time.sleep(2)
    contracts = driver.find_elements(By.CSS_SELECTOR, ".contract-row")  # 契約行のセレクタを指定

    # 契約期間が切れそうな契約を検索
    for contract in contracts:
        end_date_str = contract.find_element(By.CSS_SELECTOR, ".end-date").text  # 終了日を取得
        end_date = datetime.datetime.strptime(end_date_str, "%Y/%m/%d")
        if end_date <= (v_Today + datetime.timedelta(days=14)):
            v_ContractList.append(contract)

    # 対象契約が存在しない場合
    if not v_ContractList:
        log_message("対象契約がありません")
        driver.quit()
        exit()

    # 前回稟議を複製（ループ処理）
    for v_CurrentContract in v_ContractList:
        v_CurrentContract.click()  # 契約詳細を開く
        time.sleep(2)
        driver.find_element(By.LINK_TEXT, "申請をコピー").click()  # 申請をコピー

        # 申請タイトルを修正
        title_field = driver.find_element(By.ID, "title")
        new_title = title_field.get_attribute("value").replace("コピー：", "")
        title_field.clear()
        title_field.send_keys(new_title)

        # 契約期間を修正 (3か月後)
        start_date_str = driver.find_element(By.ID, "start_date").get_attribute("value")
        end_date_str = driver.find_element(By.ID, "end_date").get_attribute("value")

        start_date = datetime.datetime.strptime(start_date_str, "%Y/%m/%d")
        end_date = datetime.datetime.strptime(end_date_str, "%Y/%m/%d")

        v_NewStartDate = start_date + datetime.timedelta(days=90)  # 3か月後
        v_NewEndDate = end_date + datetime.timedelta(days=90)  # 3か月後

        driver.find_element(By.ID, "start_date").clear()
        driver.find_element(By.ID, "start_date").send_keys(v_NewStartDate.strftime("%Y/%m/%d"))

        driver.find_element(By.ID, "end_date").clear()
        driver.find_element(By.ID, "end_date").send_keys(v_NewEndDate.strftime("%Y/%m/%d"))

        # 下書き保存前に静止
        print("下書き保存を行う直前の処理まで完了しました。")
        input("下書き保存を行うにはEnterキーを押してください。")

        # 下書き保存
        driver.find_element(By.LINK_TEXT, "下書き保存").click()
        time.sleep(2)

except Exception as e:
    log_message(f"エラーが発生しました: {e}")
    driver.quit()
    exit()

# 処理終了
log_message("処理が完了しました")
driver.quit()