# -*- coding: utf-8 -*-
"""
全ランキング包括的テスト
app.pyのranking_optionsに定義されている全ランキング（約170件）をテスト
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scraper import OriconScraper
import time
import json
from datetime import datetime

# app.pyから全ranking_optionsを取得
RANKING_OPTIONS = {
    # ========================================
    # 保険
    # ========================================
    "自動車保険（ダイレクト型）": "_insurance",
    "自動車保険（代理店型）": "_insurance@type02",
    "自動車保険（FP推奨）": "_insurance@type03",
    "バイク保険（ダイレクト型）": "_bike",
    "バイク保険（代理店型）": "_bike@type02",
    "自転車保険": "bicycle-insurance",
    "火災保険": "fire-insurance",
    "海外旅行保険": "travel-insurance",
    "ペット保険": "_pet",
    "生命保険": "life-insurance",
    "医療保険": "medical_insurance",
    "がん保険": "cancer-insurance",
    "学資保険": "educational-insurance",
    "保険相談ショップ": "_hokenshop",
    # ========================================
    # 金融・投資
    # ========================================
    "ネット証券（顧客満足度）": "_certificate",
    "ネット証券（FP評価）": "_certificate@type02",
    "NISA（証券会社）": "_nisa",
    "NISA（銀行）": "_nisa@type02",
    "NISA（FP評価）": "_nisa@type03",
    "iDeCo（証券会社）": "ideco",
    "iDeCo（FP評価）": "ideco@type02",
    "ネット銀行": "_netbank/bank",
    "インターネットバンキング": "_netbank/banking",
    "住宅ローン": "_housingloan",
    "住宅ローン（FP評価）": "_housingloan@type02",
    "外貨預金": "foreign-currency-deposits",
    "FX（顧客満足度）": "_fx",
    "FX（FP評価）": "_fx@type02",
    "カードローン（銀行系）": "card-loan",
    "カードローン（ノンバンク）": "card-loan/nonbank",
    "クレジットカード（一般）": "credit-card/general",
    "クレジットカード（年会費無料）": "credit-card/free-annual",
    "クレジットカード（ゴールド）": "credit-card/gold-card",
    "キャッシュレス決済アプリ": "smartphone-payment",
    "暗号資産（現物取引）": "cryptocurrency/cash-transaction",
    "暗号資産（証拠金取引）": "cryptocurrency/margin-transaction",
    "ロボアドバイザー": "robo-advisor",
    # ========================================
    # 住宅・不動産
    # ========================================
    "不動産仲介 売却（マンション）": "estate-agency-sell/mansion",
    "不動産仲介 売却（戸建て）": "estate-agency-sell/kodate",
    "不動産仲介 売却（土地）": "estate-agency-sell/land",
    "不動産仲介 購入（マンション）": "estate-agency-buy/mansion",
    "不動産仲介 購入（戸建て）": "estate-agency-buy/kodate",
    "分譲マンション管理会社（首都圏）": "mansion-maintenance/syutoken",
    "分譲マンション管理会社（東海）": "mansion-maintenance/tokai",
    "分譲マンション管理会社（近畿）": "mansion-maintenance/kinki",
    "分譲マンション管理会社（九州）": "mansion-maintenance/kyusyu",
    "マンション大規模修繕": "mansion-large-repair",
    "不動産仲介 賃貸": "rental-housing",
    "賃貸情報サイト": "rental-housing/website",
    "賃貸マンション": "rental-condominiums",
    "リフォーム（フルリフォーム）": "_reform/large",
    "リフォーム（戸建て）": "_reform/kodate",
    "リフォーム（マンション）": "_reform/mansion",
    "新築分譲マンション（首都圏）": "new-condominiums/syutoken",
    "新築分譲マンション（東海）": "new-condominiums/tokai",
    "新築分譲マンション（近畿）": "new-condominiums/kinki",
    "新築分譲マンション（九州）": "new-condominiums/kyusyu",
    "ハウスメーカー（注文住宅）": "house-maker",
    "建売住宅PB（東北）": "new-ready-built-house/powerbuilder/tohoku",
    "建売住宅PB（北関東）": "new-ready-built-house/powerbuilder/kitakanto",
    "建売住宅PB（首都圏）": "new-ready-built-house/powerbuilder/syutoken",
    "建売住宅PB（東海）": "new-ready-built-house/powerbuilder/tokai",
    "建売住宅PB（近畿）": "new-ready-built-house/powerbuilder/kinki",
    "建売住宅PB（九州）": "new-ready-built-house/powerbuilder/kyusyu",
    # ========================================
    # 生活サービス
    # ========================================
    "ウォーターサーバー": "_waterserver",
    "ウォーターサーバー（浄水型）": "_waterserver/purifier",
    "家事代行サービス": "housekeeping",
    "ハウスクリーニング": "house-cleaning",
    "コインランドリー": "laundromat",
    "食材宅配（首都圏）": "food-delivery/syutoken",
    "食材宅配（東海）": "food-delivery/tokai",
    "食材宅配（近畿）": "food-delivery/kinki",
    "ミールキット（首都圏）": "food-delivery/meal-kit/syutoken",
    "ミールキット（東海）": "food-delivery/meal-kit/tokai",
    "ミールキット（近畿）": "food-delivery/meal-kit/kinki",
    "ネットスーパー": "net-super",
    "フードデリバリーサービス": "food-delivery-service",
    "ふるさと納税サイト": "hometown-tax-website",
    "トランクルーム（レンタル収納）": "trunk-room/rental-storage-space",
    "トランクルーム（コンテナ）": "trunk-room/container",
    "トランクルーム（宅配型）": "trunk-room/delivery",
    "引越し会社": "_move",
    "カーシェアリング": "carsharing",
    "レンタカー": "rent-a-car",
    "格安レンタカー": "rent-a-car/reasonable",
    "車買取会社": "_carbuyer",
    "中古車情報サイト": "used-car-sell",
    "バイク販売店": "bike-sell",
    "車検": "vehicle-inspection",
    "カーメンテナンスサービス": "car-maintenance",
    "カフェ": "cafe",
    "定額制動画配信サービス": "svod",
    "動画配信（ジャンル別：映画）": "svod/genre",
    "動画配信（ジャンル別：洋画）": "svod/genre/foreign-film",
    "動画配信（ジャンル別：国内ドラマ）": "svod/genre/japanese-drama",
    "動画配信（ジャンル別：海外ドラマ）": "svod/genre/foreign-drama",
    "動画配信（ジャンル別：韓国ドラマ）": "svod/genre/korean-drama",
    "動画配信（ジャンル別：アニメ）": "svod/genre/anime",
    "動画配信（ジャンル別：バラエティ）": "svod/genre/variety",
    "動画配信（ジャンル別：ドキュメンタリー）": "svod/genre/documentary",
    "動画配信（ジャンル別：スポーツ）": "svod/genre/sports",
    "動画配信（ジャンル別：オリジナル）": "svod/genre/original",
    "動画配信（ジャンル別：キッズ）": "svod/genre/kids",
    "子ども写真スタジオ": "kids-photo-studio",
    "電子書籍サービス": "ebook",
    "電子コミックサービス": "manga-apps",
    "マンガアプリ（オリジナル）": "manga-apps/original",
    "マンガアプリ（出版社）": "manga-apps/publisher",
    "ブランド品買取（店舗）": "brand-sell",
    "電力会社（小売）": "electricity/retailing",
    "子ども見守りGPS": "child-gps",
    # ========================================
    # 通信
    # ========================================
    "携帯キャリア": "mobile-carrier",
    "キャリア格安ブランド": "mobile-carrier/reasonable",
    "格安SIM": "mvno",
    "格安SIM（SIMのみ）": "mvno/sim",
    "格安スマホ": "mvno/sp",
    "プロバイダ": "_internet",
    "プロバイダ（北海道）": "_internet/hokkaido",
    "プロバイダ（東北）": "_internet/tohoku",
    "プロバイダ（関東）": "_internet/kanto",
    "プロバイダ（甲信越・北陸）": "_internet/koshinetsu-hokuriku",
    "プロバイダ（東海）": "_internet/tokai",
    "プロバイダ（近畿）": "_internet/kinki",
    "プロバイダ（中国）": "_internet/chugoku",
    "プロバイダ（四国）": "_internet/shikoku",
    "プロバイダ（九州・沖縄）": "_internet/kyusyu-okinawa",
    # ========================================
    # 教育（塾・受験）※juken.oricon.co.jp
    # ========================================
    "大学受験 塾・予備校（首都圏）": "_college/syutoken",
    "大学受験 塾・予備校（東海）": "_college/tokai",
    "大学受験 塾・予備校（近畿）": "_college/kinki",
    "大学受験 個別指導塾（首都圏）": "college-individual/syutoken",
    "大学受験 個別指導塾（東海）": "college-individual/tokai",
    "大学受験 個別指導塾（近畿）": "college-individual/kinki",
    "大学受験 難関大学特化型（首都圏）": "_college/elite",
    "大学受験 映像授業": "college-video",
    "高校受験 塾（北海道）": "highschool/hokkaido",
    "高校受験 塾（東北）": "highschool/tohoku",
    "高校受験 塾（北関東）": "highschool/kitakanto",
    "高校受験 塾（首都圏）": "highschool/syutoken",
    "高校受験 塾（甲信越・北陸）": "highschool/koshinetsu-hokuriku",
    "高校受験 塾（東海）": "highschool/tokai",
    "高校受験 塾（近畿）": "highschool/kinki",
    "高校受験 塾（中国・四国）": "highschool/chugoku-shikoku",
    "高校受験 塾（九州・沖縄）": "highschool/kyusyu",
    "高校受験 個別指導塾（北海道）": "highschool-individual/hokkaido",
    "高校受験 個別指導塾（東北）": "highschool-individual/tohoku",
    "高校受験 個別指導塾（北関東）": "highschool-individual/kitakanto",
    "高校受験 個別指導塾（首都圏）": "highschool-individual/syutoken",
    "高校受験 個別指導塾（甲信越・北陸）": "highschool-individual/koshinetsu-hokuriku",
    "高校受験 個別指導塾（東海）": "highschool-individual/tokai",
    "高校受験 個別指導塾（近畿）": "highschool-individual/kinki",
    "高校受験 個別指導塾（中国・四国）": "highschool-individual/chugoku-shikoku",
    "高校受験 個別指導塾（九州・沖縄）": "highschool-individual/kyusyu",
    "中学受験 塾（首都圏）": "_junior/syutoken",
    "中学受験 塾（東海）": "_junior/tokai",
    "中学受験 塾（近畿）": "_junior/kinki",
    "中学受験 個別指導塾": "_junior/individual",
    "公立中高一貫校対策 塾（首都圏）": "public-junior/syutoken",
    "公立中高一貫校対策 塾（東海）": "public-junior/tokai",
    "公立中高一貫校対策 塾（近畿）": "public-junior/kinki",
    # ========================================
    # 教育（通信・英語・資格）※juken.oricon.co.jp
    # ========================================
    "通信教育（高校生）": "online-study/highschool",
    "通信教育（中学生）": "online-study/junior-hs",
    "通信教育（小学生）": "online-study/elementary",
    "家庭教師": "tutor",
    "補習塾": "supplementary-school",
    "幼児・小学生 学習教室": "kids-school/intellectual",
    "子ども英語教室（幼児）": "kids-english/preschooler",
    "子ども英語教室（小学生）": "kids-english/grade-schooler",
    "英会話スクール": "_english",
    "オンライン英会話": "online-english",
    "通信講座（FP）": "cc/fp",
    "通信講座（医療事務）": "cc/mo",
    "通信講座（宅建）": "cc/takken",
    "通信講座（簿記）": "cc/bookkeeping",
    "通信講座（TOEIC）": "cc/toeic",
    "通信講座（社会保険労務士）": "cc/labor-and-social-security",
    "通信講座（ケアマネジャー）": "cc/care-manager",
    "通信講座（公務員）": "cc/public-officer",
    "通信講座（ITパスポート）": "cc/it-certification",
    "資格スクール（FP）": "license/fp",
    "資格スクール（宅建）": "license/takken",
    "資格スクール（簿記）": "license/bookkeeping",
    "資格スクール（社会保険労務士）": "license/labor-and-social-security",
    # ========================================
    # スポーツ・フィットネス
    # ========================================
    "キッズスイミングスクール（幼児）": "kids-swimming/preschooler",
    "キッズスイミングスクール（小学生）": "kids-swimming/grade-schooler",
    "フィットネスクラブ": "_fitness",
    "24時間ジム": "_fitness/24hours",
    "パーソナルトレーニング": "_fitness/service",
    # ========================================
    # 転職・人材 ※career.oricon.co.jp
    # ========================================
    "就活サイト": "new-graduates-hiring-website",
    "逆求人型就活サービス": "reversed-job-offer",
    "アルバイト情報サイト": "arbeit",
    "転職サイト": "job-change",
    "転職サイト（女性）": "job-change_woman",
    "転職スカウトサービス": "job-change_scout",
    "転職エージェント": "_agent",
    "看護師転職": "_agent_nurse",
    "介護転職": "_agent_nursing",
    "ハイクラス・ミドルクラス転職": "_agent_hi-and-middle-class",
    "派遣会社": "_staffing",
    "工場・製造業派遣": "_staffing_manufacture",
    "派遣求人サイト": "temp-staff",
    "求人情報サービス": "employment",
    # ========================================
    # トラベル
    # ========================================
    "旅行予約サイト（国内）": "bargain-hotels-website",
    "旅行予約サイト（海外）": "bargain-airline-website",
    "ツアー比較サイト": "tours-website",
    # ========================================
    # 美容・ウエディング
    # ========================================
    "ブライダルエステ": "_esthe/bridal",
    "フェイシャルエステ": "_esthe/facial",
    "痩身・ボディエステ": "_esthe/slim",
    "サロン検索予約サイト": "salon-website",
    "ハウスウエディング": "wedding-produce",
    "結婚相談所": "_marriage",
    # ========================================
    # 小売・レジャー
    # ========================================
    "家電量販店": "electronics-retail-store",
    "ドラッグストア": "drug-store",
    "映画館": "movie-theater",
    "カラオケボックス": "karaoke",
    "テーマパーク": "theme-park",
}

def test_all_rankings():
    """全ランキングの部門・評価項目検出をテスト"""
    print("=" * 80)
    print(f"全ランキング包括的テスト - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"テスト対象: {len(RANKING_OPTIONS)}件")
    print("=" * 80)
    print()

    results = []
    problems = []
    errors = []

    for i, (name, slug) in enumerate(RANKING_OPTIONS.items(), 1):
        print(f"[{i:3d}/{len(RANKING_OPTIONS)}] テスト中: {name} ({slug})")

        try:
            scraper = OriconScraper(slug, name)
            main_url = f"{scraper.BASE_URL}/{scraper.url_prefix}/"

            if scraper.subpath:
                main_url = f"{scraper.BASE_URL}/{scraper.url_prefix}/{scraper.subpath}/"

            # 部門検出
            departments = scraper._discover_departments(main_url)
            dept_count = len(departments)

            # 評価項目検出
            items = scraper._discover_evaluation_items(main_url)
            item_count = len(items)

            results.append({
                "name": name,
                "slug": slug,
                "url": main_url,
                "departments": dept_count,
                "items": item_count,
                "dept_names": list(departments.values())[:5],
                "item_names": list(items.values())[:5],
            })

            status = "OK" if (dept_count > 0 or item_count > 0) else "WARN"
            print(f"        -> 部門: {dept_count}件, 評価項目: {item_count}件 [{status}]")

            if dept_count == 0 and item_count == 0:
                problems.append({
                    "name": name,
                    "slug": slug,
                    "url": main_url,
                    "issue": "部門・評価項目ともに検出なし"
                })

        except Exception as e:
            error_msg = str(e)[:80]
            print(f"        -> エラー: {error_msg}")
            errors.append({
                "name": name,
                "slug": slug,
                "issue": error_msg
            })

        time.sleep(0.3)  # サーバー負荷軽減

    # サマリー出力
    print()
    print("=" * 80)
    print("テスト結果サマリー")
    print("=" * 80)
    print()

    print(f"テスト総数: {len(RANKING_OPTIONS)}")
    print(f"成功: {len(results)}")
    print(f"警告（部門・項目0件）: {len(problems)}")
    print(f"エラー: {len(errors)}")
    print()

    # 統計
    dept_only = sum(1 for r in results if r["departments"] > 0 and r["items"] == 0)
    item_only = sum(1 for r in results if r["departments"] == 0 and r["items"] > 0)
    both = sum(1 for r in results if r["departments"] > 0 and r["items"] > 0)
    neither = sum(1 for r in results if r["departments"] == 0 and r["items"] == 0)

    print("検出状況:")
    print(f"  - 部門+評価項目あり: {both}件")
    print(f"  - 部門のみ: {dept_only}件")
    print(f"  - 評価項目のみ: {item_only}件")
    print(f"  - どちらもなし: {neither}件")
    print()

    # 問題詳細
    if problems:
        print("=" * 80)
        print(f"警告の詳細 ({len(problems)}件)")
        print("=" * 80)
        for p in problems:
            print(f"  - {p['name']} ({p['slug']})")
            print(f"    URL: {p['url']}")
        print()

    if errors:
        print("=" * 80)
        print(f"エラーの詳細 ({len(errors)}件)")
        print("=" * 80)
        for e in errors:
            print(f"  - {e['name']} ({e['slug']}): {e['issue']}")
        print()

    # 結果をJSONファイルに保存
    report = {
        "timestamp": datetime.now().isoformat(),
        "total": len(RANKING_OPTIONS),
        "success": len(results),
        "warnings": len(problems),
        "errors": len(errors),
        "statistics": {
            "both_dept_and_items": both,
            "dept_only": dept_only,
            "items_only": item_only,
            "neither": neither,
        },
        "results": results,
        "problems": problems,
        "errors": errors,
    }

    report_path = os.path.join(os.path.dirname(__file__), "test_report_comprehensive.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"詳細レポート: {report_path}")

    return results, problems, errors

if __name__ == "__main__":
    results, problems, errors = test_all_rankings()
