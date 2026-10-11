import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_excel_report():
    wb = openpyxl.Workbook()
    # デフォルトシートを削除して新規作成
    wb.remove(wb.active)

    # カラーパレット（洗練されたモノトーンベース）
    C_HEADER_BG = "262626"       # 最上部タイトル・主見出し (ダークチャコール)
    C_HEADER_FG = "FFFFFF"       # 白文字
    C_SECTION_BG = "404040"      # セクション見出し (濃いグレー)
    C_SECTION_FG = "FFFFFF"      # 白文字
    C_SUBSECTION_BG = "595959"   # テーブルヘッダー (ミディアムダークグレー)
    C_SUBSECTION_FG = "FFFFFF"   # 白文字
    C_TABLE_HEADER_BG = "E6E6E6" # サブテーブルヘッダー (ライトグレー)
    C_TABLE_HEADER_FG = "1A1A1A" # 濃い黒文字
    C_ZEBRA_LIGHT = "F9F9F9"     # 偶数行 (超薄いグレー)
    C_ZEBRA_WHITE = "FFFFFF"     # 奇数行 (白)
    C_HIGHLIGHT_BG = "EDEDED"    # 注目枠・推奨枠背景
    C_BORDER = "BFBFBF"          # 通常罫線 (薄いグレー)
    C_BORDER_DARK = "404040"     # 強調罫線

    # フォント定義 (Meiryo UI を徹底)
    FONT_FAMILY = "Meiryo UI"
    f_title = Font(name=FONT_FAMILY, size=14, bold=True, color="FFFFFF")
    f_section = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    f_th_dark = Font(name=FONT_FAMILY, size=9.5, bold=True, color="FFFFFF")
    f_th_light = Font(name=FONT_FAMILY, size=9.5, bold=True, color="1A1A1A")
    f_bold = Font(name=FONT_FAMILY, size=9, bold=True, color="000000")
    f_normal = Font(name=FONT_FAMILY, size=9, bold=False, color="1A1A1A")
    f_normal_gray = Font(name=FONT_FAMILY, size=8.5, bold=False, color="555555")
    f_badge_ebs = Font(name=FONT_FAMILY, size=9, bold=True, color="000000")
    f_badge_efs = Font(name=FONT_FAMILY, size=9, bold=False, color="444444")

    # 罫線定義
    thin_side = Side(border_style="thin", color=C_BORDER)
    thick_bottom_side = Side(border_style="medium", color=C_BORDER_DARK)
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_th_dark = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_bottom_side)
    border_th_light = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # 塗りつぶし定義
    fill_title = PatternFill(start_color=C_HEADER_BG, end_color=C_HEADER_BG, fill_type="solid")
    fill_section = PatternFill(start_color=C_SECTION_BG, end_color=C_SECTION_BG, fill_type="solid")
    fill_th_dark = PatternFill(start_color=C_SUBSECTION_BG, end_color=C_SUBSECTION_BG, fill_type="solid")
    fill_th_light = PatternFill(start_color=C_TABLE_HEADER_BG, end_color=C_TABLE_HEADER_BG, fill_type="solid")
    fill_zebra_light = PatternFill(start_color=C_ZEBRA_LIGHT, end_color=C_ZEBRA_LIGHT, fill_type="solid")
    fill_zebra_white = PatternFill(start_color=C_ZEBRA_WHITE, end_color=C_ZEBRA_WHITE, fill_type="solid")
    fill_highlight = PatternFill(start_color=C_HIGHLIGHT_BG, end_color=C_HIGHLIGHT_BG, fill_type="solid")

    # 配置定義
    al_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
    al_right = Alignment(horizontal="right", vertical="center", wrap_text=False)

    # =========================================================================
    # Sheet 1: 01_結論と選定サマリー
    # =========================================================================
    ws1 = wb.create_sheet(title="01_結論と選定サマリー")
    ws1.views.sheetView[0].showGridLines = True

    # タイトル
    ws1.merge_cells("A1:G1")
    cell = ws1["A1"]
    cell.value = "【結論・設計提言】EC2バッチサーバにおけるディスク選定（Amazon EBS vs Amazon EFS）"
    cell.font = f_title
    cell.fill = fill_title
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[1].height = 36

    # サブタイトル / ワークロード概要
    ws1.merge_cells("A2:G2")
    ws1["A2"].value = "対象環境: IT環境(c7a.2xlarge, 10JVM並列でCPU90%超) → 本番環境(32vCPU, 58JVM並列, 32マルチ実行) / 用途: ログ出力領域 & 一時ファイル(Temp)領域"
    ws1["A2"].font = f_normal_gray
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[2].height = 20

    # セクション1: 結論サマリー
    ws1.merge_cells("A4:G4")
    ws1["A4"].value = "1. 設計判断の結論：Amazon EBS（汎用SSD: gp3）の圧倒的一択"
    ws1["A4"].font = f_section
    ws1["A4"].fill = fill_section
    ws1["A4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[4].height = 24

    ws1.merge_cells("A5:G6")
    ws1["A5"].value = (
        "【診断結論】本件ワークロード（58JVM並列・高頻度ログ書き出し・一時ファイル生成）において、Amazon EFSの採用は極めて危険であり、EBS（gp3）の採用を強く推奨します。\n"
        "EFSを採用した場合、NFSv4.1経由のネットワーク遅延および3AZ同期合意形成により、高頻度のログ追記(write)や一時ファイルの作成・削除(create/delete)でJVMスレッドが大量にI/O待ち(BLOCKED)に陥り、CPUが遊んでいるのにバッチ処理時間が数倍〜十数倍に激増します。"
    )
    ws1["A5"].font = f_bold
    ws1["A5"].fill = fill_highlight
    ws1["A5"].alignment = al_left_top
    ws1.row_dimensions[5].height = 22
    ws1.row_dimensions[6].height = 22

    # 決定打となる3つの理由テーブル
    ws1.merge_cells("A8:G8")
    ws1["A8"].value = "2. EBSを選ぶべき＆EFSを避けるべき決定的な3大理由"
    ws1["A8"].font = f_section
    ws1["A8"].fill = fill_section
    ws1["A8"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[8].height = 24

    headers_3reasons = ["No", "評価軸", "Amazon EBS (gp3) の挙動", "Amazon EFS の挙動", "本件バッチ処理 (58JVM) への影響度・リスク"]
    ws1.cell(row=9, column=1, value="No").alignment = al_center
    ws1.cell(row=9, column=2, value="選定の観点").alignment = al_center
    ws1.cell(row=9, column=3, value="Amazon EBS (推奨)").alignment = al_center
    ws1.cell(row=9, column=4, value="Amazon EFS (非推奨)").alignment = al_center
    ws1.merge_cells("E9:G9")
    ws1.cell(row=9, column=5, value="本件ワークロード(58JVM並列)におけるリスク・影響度").alignment = al_center
    ws1.row_dimensions[9].height = 24

    for col in range(1, 8):
        c = ws1.cell(row=9, column=col)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.border = border_th_dark

    reasons_data = [
        (
            "1",
            "I/Oレイテンシ\n(応答速度の格差)",
            "【0.5ms 〜 1ms 未満】\nローカルNVMe直結同等のブロックアクセス。OSページキャッシュもフル活用され、ログ書き込みはマイクロ秒で即時復帰。",
            "【2ms 〜 15ms 以上】\nネットワーク経由のNFS RPC往復＋3AZ同期レプリケーションの合意形成が必要。小サイズ書き込みが桁違いに遅い。",
            "【極めて重大（処理遅延）】\n58並列のJVMがログ追記(append/flush)を行うたびにミリ秒単位の待ち時間が発生。スレッドが滞留し、バッチ完了時間が破綻する。"
        ),
        (
            "2",
            "メタデータ処理速度\n(ファイル作成・削除)",
            "【極小（OSメモリ内処理）】\nXFS/ext4のinode管理はEC2ローカルカーネルが行うため、createTempFileやdeleteは数マイクロ秒で完了。",
            "【強烈なオーバーヘッド】\nファイルの作成(CREATE)・削除(REMOVE)・属性取得(GETATTR)ごとに分散クラスタと通信。並列操作で直列化・競合が発生。",
            "【致命的（I/Oボトルネック）】\nバッチ処理が一時ファイルを頻繁に生成・破棄すると、EFSのメタデータ層がボトルネック化し、I/O Waitが急上昇する。"
        ),
        (
            "3",
            "共有機能の必要性\nと費用対効果",
            "【単一EC2専用で最適】\n複数サーバ共有はできないが、今回は1台のEC2バッチサーバで完結するため全く問題なし。コストは東京で $0.08/GB-月 と安価。",
            "【不要な機能に高コスト】\n数百台のEC2から同時マウントするための分散ファイルシステム。単一EC2ではメリットが皆無な上、コストは $0.30/GB-月 と約3.7倍高価。",
            "【無駄なコスト＆性能劣化】\n全く使わない「マルチAZ共有機能」のために、高い利用料を払い、性能低下と運用リスクを抱え込むことになる。"
        )
    ]

    r_idx = 10
    for r_num, view, ebs_txt, efs_txt, impact in reasons_data:
        ws1.row_dimensions[r_idx].height = 54
        ws1.cell(row=r_idx, column=1, value=r_num).alignment = al_center
        ws1.cell(row=r_idx, column=2, value=view).alignment = al_center
        ws1.cell(row=r_idx, column=3, value=ebs_txt).alignment = al_left
        ws1.cell(row=r_idx, column=4, value=efs_txt).alignment = al_left
        ws1.merge_cells(start_row=r_idx, start_column=5, end_row=r_idx, end_column=7)
        ws1.cell(row=r_idx, column=5, value=impact).alignment = al_left

        for c_idx in range(1, 8):
            c = ws1.cell(row=r_idx, column=c_idx)
            c.font = f_normal
            c.fill = fill_zebra_white if r_num in ["1", "3"] else fill_zebra_light
            c.border = border_cell
        r_idx += 1

    # セクション3: 小学生にもわかる直感的な例え話
    ws1.merge_cells(f"A{r_idx+1}:G{r_idx+1}")
    ws1[f"A{r_idx+1}"].value = "3. 小学生にもわかる！直感的なたとえ話で見るEBSとEFSの違い"
    ws1[f"A{r_idx+1}"].font = f_section
    ws1[f"A{r_idx+1}"].fill = fill_section
    ws1[f"A{r_idx+1}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[r_idx+1].height = 24

    kids_rows = [
        ("ストレージ", "Amazon EBS", "Amazon EFS"),
        ("日常のたとえ", "【自分の勉強机の引き出し・手元のノート】", "【学校の職員室にある、全校生徒共有の大きな本棚】"),
        ("特徴と使い方", "自分の机のすぐ下にあるので、手を伸ばせば0.1秒で届きます。\n誰にも邪魔されず、計算のメモ用紙（一時ファイル）を殴り書きしてはポイポイ捨てられます。", "どの教室（どのサーバ）からでも同じ本が見られる便利な本棚です。\nただし、メモを1行書くたびに職員室まで走って行かなければなりません。"),
        ("先生の確認\n(3AZレプリカ)", "自分の引き出しなので、先生の許可はいりません。\nどんどん自分のペースで宿題（バッチ計算）を進められます。", "職員室には3人の先生がいて、あなたがメモを挟むたびに3人全員が「よし記録した！」とハンコを押すまで教室に帰れません。"),
        ("58人が同時に\nメモを書いたら？", "58人の生徒が自分の机で猛烈なスピードで計算します。\n32本の腕（32vCPU）をフル回転させて、あっという間に宿題が終わります！", "58人が一斉に職員室へダッシュ！廊下は大渋滞になり、職員室の前は長蛇の列。\n「返事が来ないから次の計算ができないよ…」と机の前でボーッと待ちぼうけ。"),
        ("子どもの結論", "◎ 自分の机で1人で勉強するなら、自分の引き出し（EBS）を使うのが当たり前！\nわざわざ職員室の本棚（EFS）に走りに行く必要はゼロ！", "× 遠くて時間がかかってヘトヘトになるので、計算用紙やメモ帳には絶対に使っちゃダメ！")
    ]

    r_kids = r_idx + 2
    ws1.cell(row=r_kids, column=1, value="比較項目").alignment = al_center
    ws1.merge_cells(start_row=r_kids, start_column=2, end_row=r_kids, end_column=4)
    ws1.cell(row=r_kids, column=2, value="Amazon EBS (自分の机の引き出し)").alignment = al_center
    ws1.merge_cells(start_row=r_kids, start_column=5, end_row=r_kids, end_column=7)
    ws1.cell(row=r_kids, column=5, value="Amazon EFS (職員室の共有本棚)").alignment = al_center
    ws1.row_dimensions[r_kids].height = 24

    for col in range(1, 8):
        c = ws1.cell(row=r_kids, column=col)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.border = border_th_dark

    r_kids += 1
    for item, ebs_val, efs_val in kids_rows[1:]:
        ws1.row_dimensions[r_kids].height = 42 if item != "特徴と使い方" else 46
        ws1.cell(row=r_kids, column=1, value=item).alignment = al_center
        ws1.merge_cells(start_row=r_kids, start_column=2, end_row=r_kids, end_column=4)
        ws1.cell(row=r_kids, column=2, value=ebs_val).alignment = al_left
        ws1.merge_cells(start_row=r_kids, start_column=5, end_row=r_kids, end_column=7)
        ws1.cell(row=r_kids, column=5, value=efs_val).alignment = al_left

        for c_idx in range(1, 8):
            c = ws1.cell(row=r_kids, column=c_idx)
            c.font = f_normal
            c.fill = fill_zebra_white if r_kids % 2 == 1 else fill_zebra_light
            c.border = border_cell
        r_kids += 1

    # =========================================================================
    # Sheet 2: 02_EBS vs EFS 総合比較表
    # =========================================================================
    ws2 = wb.create_sheet(title="02_総合詳細比較マトリクス")
    ws2.views.sheetView[0].showGridLines = True

    # タイトル
    ws2.merge_cells("A1:G1")
    ws2["A1"].value = "【技術仕様・性能・コスト】Amazon EBS vs Amazon EFS 詳細比較マトリクス"
    ws2["A1"].font = f_title
    ws2["A1"].fill = fill_title
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[1].height = 36

    # テーブルヘッダー
    ws2.row_dimensions[3].height = 28
    matrix_headers = ["大分類", "詳細項目", "Amazon EBS (gp3)\n【推奨】", "Amazon EBS (io2 Block Express)\n【参考: 超高性能】", "Amazon EFS (Elastic/GP)\n【非推奨】", "Amazon EFS (One Zone)\n【非推奨】", "技術的差異・本件ワークロードへの適用評価"]
    for idx, h_text in enumerate(matrix_headers, start=1):
        c = ws2.cell(row=3, column=idx, value=h_text)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.alignment = al_center
        c.border = border_th_dark

    matrix_data = [
        ("基本構造", "ストレージ種別", "ブロックストレージ\n(Raw Block Device)", "ブロックストレージ\n(SANアーキテクチャ)", "分散ファイルストレージ\n(NFSv4.1)", "分散ファイルストレージ\n(NFSv4.1 - 1AZ版)", "EBSはOSが直接ファイルシステムを管理。EFSはNFSプロトコル通信。"),
        ("基本構造", "プロトコル / 接続層", "NVMe over PCIe\n(Nitroハードウェア仮想化)", "NVMe over PCIe\n(Nitroハードウェア仮想化)", "NFSv4.1 over TCP/IP\n(ポート 2049)", "NFSv4.1 over TCP/IP\n(ポート 2049)", "EBSはPCIeバス直結デバイス。EFSはネットワークプロトコルスタックを経由。"),
        ("基本構造", "耐久性・配置スコープ", "単一AZ内レプリケーション\n(99.8% 〜 99.9% 耐久性)", "単一AZ内レプリケーション\n(99.999% 最高耐久性)", "3つのAZにまたがる同期冗長\n(99.999999999% 耐久性)", "単一AZ内に限定\n(99.9999999% 耐久性)", "ログや一時ファイルに3AZ冗長は過剰品質。高耐久性の代償としてEFSは遅い。"),
        ("基本構造", "ファイルシステム", "XFS / ext4 等を任意フォーマット\n(OSローカル管理)", "XFS / ext4 等を任意フォーマット\n(OSローカル管理)", "NFSv4 準拠マネージドFS\n(POSIX互換だが制約あり)", "NFSv4 準拠マネージドFS\n(POSIX互換だが制約あり)", "XFSの最適化されたアロケーショングループが使えるEBSがバッチ並列に最適。"),
        ("性能指標", "アクセスレイテンシ", "【0.5ms 〜 1.0ms 未満】\n極めて高速・安定", "【サブミリ秒 (0.2〜0.5ms)】\n極限の低遅延", "【2.0ms 〜 15.0ms+】\nネットワーク＆3AZ合意遅延", "【1.5ms 〜 10.0ms+】\n1AZだがNFS通信遅延あり", "EBSが10倍以上高速。高頻度のログ書き込み・一時ファイルI/Oで決定的な差となる。"),
        ("性能指標", "メタデータ操作性能\n(create, delete, stat)", "極小 (数マイクロ秒)\nカーネルメモリ上で即完了", "極小 (数マイクロ秒)\nカーネルメモリ上で即完了", "低速 (数ミリ秒〜十数ミリ秒)\n分散合意RPCが発生", "低速 (数ミリ秒〜十数ミリ秒)\nRPC往復オーバーヘッドあり", "58JVMが同時に一時ファイル生成・削除を行うと、EFSはメタデータ競合で詰まる。"),
        ("性能指標", "ベースラインIOPS", "3,000 IOPS (容量無関係)\n最大16,000 IOPSまで拡張可", "最大 256,000 IOPS\n(容量比 1,000 IOPS/GB)", "ワークロードに応じて自動スケール\n(クラスタ全体で数万IOPS)", "ワークロードに応じて自動スケール\n(クラスタ全体で数万IOPS)", "単一EC2から高頻度I/Oを引き出す場合、EBSの専有IOPSの方が遥かに安定。"),
        ("性能指標", "スループット", "125 MB/s (無償枠)\n最大 1,000 MB/s までプロビジョニング可", "最大 4,000 MB/s", "読み取り最大 3+ GiB/s\n書き込み最大 1+ GiB/s (Elastic)", "読み取り最大 3+ GiB/s\n書き込み最大 1+ GiB/s (Elastic)", "gp3は設定した帯域が100%確実に保証される。一時ファイル大容量転送も余裕。"),
        ("ネットワーク", "インスタンスネットワーク帯域", "専用の「EBS最適化帯域」\n(一般ENIと完全分離)", "専用の「EBS最適化帯域」\n(一般ENIと完全分離)", "インスタンスの一般ENI帯域\n(業務API/DB通信と共有)", "インスタンスの一般ENI帯域\n(業務API/DB通信と共有)", "c7a.2xlarge(最大12.5Gbps)において、EFSはDB通信と帯域を取り合って悪影響。"),
        ("同時接続性", "マウント可能EC2台数", "原則1台\n(※Multi-Attachは特定条件)", "原則1台\n(※Multi-Attachは特定条件)", "数千台から同時マウント可能", "数千台から同時マウント可能", "今回は「単一EC2バッチサーバ」で動かすため、EFSの複数共有機能は不要。"),
        ("運用管理", "容量サイジング", "事前プロビジョニング\n(必要に応じて動的拡張可能)", "事前プロビジョニング\n(必要に応じて動的拡張可能)", "完全自動拡張\n(ペタバイトまで意識不要)", "完全自動拡張\n(ペタバイトまで意識不要)", "一時ファイル領域は容量予測が容易であり、EBSのサイズ管理は全く苦にならない。"),
        ("コスト比較", "ストレージ容量単価\n(東京リージョン)", "$0.08 / GB-月\n(gp2より20%安価)", "$0.142 / GB-月\n+ IOPS課金 ($0.0705/IOPS)", "$0.30 / GB-月 (Standard)\n(EBS gp3 の約3.75倍)", "$0.16 / GB-月 (One Zone)\n(EBS gp3 の約2.0倍)", "EFSは単純な容量単価だけでもEBSの約3.75倍高価。大量データ保持はコスト大。"),
        ("コスト比較", "スループット/IOPS課金", "ベースライン(3,000 IOPS / 125MB/s)無料\n追加時のみ課金", "プロビジョンドIOPS課金あり", "Elastic Throughput課金:\n読取 $0.03/GB, 書込 $0.06/GB", "Elastic Throughput課金:\n読取 $0.03/GB, 書込 $0.06/GB", "一時ファイルやログを大量に出し入れすると、EFSはスループット課金が跳ね上がる。"),
        ("適合性判定", "バッチログ出力領域", "【◎ 最適】\n追記が低遅延でJVMを止めない", "【○ 優秀】\n高価だが性能は最高峰", "【× 不適】\nフラッシュ遅延でスレッド滞留", "【× 不適】\nNFSレイテンシでスレッド滞留", "ログ出力は高頻度小サイズ書き込み。EBS gp3がコスト・性能ともにベスト。"),
        ("適合性判定", "バッチ一時(Temp)領域", "【◎ 最適】\n高速ファイル作成・破棄・シーク", "【◎ 最適】\n超高速だがコスト過剰", "【×× 致命的】\nメタデータRPCで激しい遅延", "【×× 致命的】\nメタデータRPCで激しい遅延", "一時ファイル生成(createTempFile)を58並列で行うとEFSは破綻する。")
    ]

    r_idx = 4
    for cat, item, ebs_gp3, ebs_io2, efs_std, efs_oz, desc in matrix_data:
        ws2.row_dimensions[r_idx].height = 36 if "ファイル" not in item else 40
        ws2.cell(row=r_idx, column=1, value=cat).alignment = al_center
        ws2.cell(row=r_idx, column=2, value=item).alignment = al_left
        ws2.cell(row=r_idx, column=3, value=ebs_gp3).alignment = al_left
        ws2.cell(row=r_idx, column=4, value=ebs_io2).alignment = al_left
        ws2.cell(row=r_idx, column=5, value=efs_std).alignment = al_left
        ws2.cell(row=r_idx, column=6, value=efs_oz).alignment = al_left
        ws2.cell(row=r_idx, column=7, value=desc).alignment = al_left

        is_even = (r_idx % 2 == 0)
        row_fill = fill_zebra_white if is_even else fill_zebra_light

        for c_idx in range(1, 8):
            c = ws2.cell(row=r_idx, column=c_idx)
            c.font = f_normal
            c.fill = row_fill
            c.border = border_cell
            if c_idx == 3: # 推奨EBS gp3列の文字を太字に
                c.font = f_bold
        r_idx += 1

    # =========================================================================
    # Sheet 3: 03_動作原理と歴史的背景
    # =========================================================================
    ws3 = wb.create_sheet(title="03_動作原理と歴史的背景")
    ws3.views.sheetView[0].showGridLines = True

    # タイトル
    ws3.merge_cells("A1:G1")
    ws3["A1"].value = "【動作原理と歴史的背景】なぜEBSとEFSが生まれ、なぜ内部挙動が根本的に異なるのか"
    ws3["A1"].font = f_title
    ws3["A1"].fill = fill_title
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws3.row_dimensions[1].height = 36

    # セクション1: 登場の歴史と背景
    ws3.merge_cells("A3:G3")
    ws3["A3"].value = "1. クラウドストレージの登場の歴史と背景（必要になった経緯）"
    ws3["A3"].font = f_section
    ws3["A3"].fill = fill_section
    ws3["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws3.row_dimensions[3].height = 24

    history_headers = ["年代・節目", "当時の課題と顧客の悲鳴", "AWSがリリースした技術革新", "ストレージの性質と設計思想", "本件バッチ処理への教訓"]
    ws3.cell(row=4, column=1, value="年代・節目").alignment = al_center
    ws3.cell(row=4, column=2, value="当時の課題と顧客の悲鳴").alignment = al_center
    ws3.cell(row=4, column=3, value="AWSが投入した技術と進化").alignment = al_center
    ws3.merge_cells("D4:E4")
    ws3.cell(row=4, column=4, value="設計思想とアーキテクチャの狙い").alignment = al_center
    ws3.merge_cells("F4:G4")
    ws3.cell(row=4, column=6, value="本件バッチへの教訓・示唆").alignment = al_center
    ws3.row_dimensions[4].height = 24

    for col in range(1, 8):
        c = ws3.cell(row=4, column=col)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.border = border_th_dark

    history_data = [
        (
            "2006年\nクラウド黎明期",
            "【インスタンスストアの揮発性問題】\nEC2登場当時、ストレージは内蔵揮発性ディスクのみ。インスタンスを停止(Stop)すると全データが物理的に消滅した。",
            "・Amazon S3 (オブジェクトストレージ)\n・EC2インスタンスストア (Ephemeral Disk)",
            "「OSが消えても困らないステートレスな処理」しかできず、データベースや基幹バッチ処理の移行が不可能だった。",
            "ディスク永続化と独立したI/Oチャネルの重要性が認識された原点。"
        ),
        (
            "2008年〜2014年\nAmazon EBS誕生\nとSSD化",
            "【OS停止してもデータが消えないディスク】\nオンプレミスの外付けHDDやSANのように、インスタンスの停止・再起動と独立して永続化できるディスクが切望された。",
            "・2008年: Amazon EBS (Magnetic) 誕生\n・2012年: io1 (プロビジョンドIOPS SSD)\n・2014年: gp2 (汎用SSD・バーストクレジット型)",
            "「1台のEC2に仮想ケーブルで外付けディスクを直結する」思想。徹底的にブロックレベルの低遅延・安定性を追求。",
            "単一サーバで動かすバッチ処理の作業ディスクとして、EBSが完全に標準となった。"
        ),
        (
            "2015年〜2016年\nAmazon EFS誕生\n(共有FSの需要)",
            "【複数サーバから同時に同じファイルを読みたい】\nWebサーバ複数台構成やCMS(WordPress)移行時、画像や設定ファイルを全台でリアルタイム共有する手段がAWSになかった。",
            "・2015年: Amazon EFS プレビュー\n・2016年: Amazon EFS 一般提供開始 (NFSv4.1)",
            "「容量無制限、数千台のEC2から同時マウント、3AZ冗長で絶対に消えないファイル共有」。速度よりも『共有性と可用性』を最優先。",
            "EFSはファイル共有用であり、バッチ処理の高速な一時計算領域用ではないことが設計思想から明白。"
        ),
        (
            "2020年〜最新\nEBS gp3 と\nEFS Elastic化",
            "【容量と性能の分離、スパイク負荷対応】\n・EBS: 容量を増やさないとIOPSが上がらないgp2の不満\n・EFS: スループット事前設定が難しく枯渇する問題",
            "・2020年: EBS gp3 (ベース3,000 IOPS標準付与)\n・2020年: io2 Block Express (SAN同等のサブミリ秒)\n・2022年: EFS Elastic Throughput (従量自動スケール)",
            "・EBS gp3: 小容量でも最初から3,000 IOPSが使え、コストと性能が劇的に改善。\n・EFS: スループットは自動化されたが、NFS RPCの物理レイテンシは残る。",
            "現在のgp3はバッチサーバの一時領域として最もコストパフォーマンスと性能が高い完成形となった。"
        )
    ]

    r_hist = 5
    for era, pain, aws_sol, phil, lesson in history_data:
        ws3.row_dimensions[r_hist].height = 54
        ws3.cell(row=r_hist, column=1, value=era).alignment = al_center
        ws3.cell(row=r_hist, column=2, value=pain).alignment = al_left
        ws3.cell(row=r_hist, column=3, value=aws_sol).alignment = al_left
        ws3.merge_cells(start_row=r_hist, start_column=4, end_row=r_hist, end_column=5)
        ws3.cell(row=r_hist, column=4, value=phil).alignment = al_left
        ws3.merge_cells(start_row=r_hist, start_column=6, end_row=r_hist, end_column=7)
        ws3.cell(row=r_hist, column=6, value=lesson).alignment = al_left

        for c_idx in range(1, 8):
            c = ws3.cell(row=r_hist, column=c_idx)
            c.font = f_normal
            c.fill = fill_zebra_white if r_hist % 2 == 1 else fill_zebra_light
            c.border = border_cell
        r_hist += 1

    # セクション2: 動作原理の深掘り比較
    ws3.merge_cells(f"A{r_hist+1}:G{r_hist+1}")
    ws3[f"A{r_hist+1}"].value = "2. 動作原理・OSレイヤー・通信フローの徹底比較"
    ws3[f"A{r_hist+1}"].font = f_section
    ws3[f"A{r_hist+1}"].fill = fill_section
    ws3[f"A{r_hist+1}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws3.row_dimensions[r_hist+1].height = 24

    arch_headers = ["レイヤー / 項目", "Amazon EBS (gp3) の内部挙動", "Amazon EFS (NFSv4.1) の内部挙動", "本件バッチ処理における技術的決定打"]
    r_arch = r_hist + 2
    ws3.cell(row=r_arch, column=1, value="内部アーキテクチャ層").alignment = al_center
    ws3.merge_cells(start_row=r_arch, start_column=2, end_row=r_arch, end_column=3)
    ws3.cell(row=r_arch, column=2, value="Amazon EBS (ローカルブロック層直結)").alignment = al_center
    ws3.merge_cells(start_row=r_arch, start_column=4, end_row=r_arch, end_column=5)
    ws3.cell(row=r_arch, column=4, value="Amazon EFS (分散NFSネットワーク層)").alignment = al_center
    ws3.merge_cells(start_row=r_arch, start_column=6, end_row=r_arch, end_column=7)
    ws3.cell(row=r_arch, column=6, value="バッチ実行時の技術的インサイト").alignment = al_center
    ws3.row_dimensions[r_arch].height = 24

    for col in range(1, 8):
        c = ws3.cell(row=r_arch, column=col)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.border = border_th_dark

    arch_data = [
        (
            "OSファイルシステム\n(VFS / FS Driver)",
            "EC2上のカーネル（XFS/ext4）が完全に統括。\ninodeテーブル、ディレクトリ構造、アロケーションマップはすべてEC2の手元で管理。",
            "カーネルのNFSクライアントモジュールが担当。\nファイル操作はすべてリモートNFSサーバへのRPC要求に変換されてネットワークへ送出される。",
            "EBSはOSのローカルファイルシステムがフルに機能するため、58並列の書き込みもXFSがメモリ上で軽快に裁く。"
        ),
        (
            "OSページキャッシュ\n(バッファキャッシュ)",
            "【ライトバックキャッシュが完全有効】\nJVMがwrite()を呼んだ瞬間、OSの空きRAM上にデータが乗って即座に応答。実ディスク書き込みは非同期にバックグラウンド実行。",
            "【Close-to-Open 一貫性の制約】\nNFSの仕様上、ファイルを閉じる(close)際や整合性確認時にリモートフラッシュが強制され、メモリキャッシュの恩恵が激減する。",
            "バッチ処理が一時ファイルを書いてすぐ破棄する場合、EBSなら「ほぼRAMの超高速性能」で動くが、EFSは都度NFS通信が発生。"
        ),
        (
            "I/Oスタックと\nドライバ層",
            "Linux Block Layer ──> NVMeドライバ ──> PCIeバス ──> Nitroチップ。\n最適化されたブロックデバイスドライバで最短パスを実行。",
            "VFS ──> NFS Client ──> TCP/IP Stack ──> ENA(ネットワークカード) ──> VPC。\nプロトコルスタックが深く、CPU処理とシリアライズ負荷が発生。",
            "EBSはAWS Nitroシステムによりハードウェアでオフロードされるため、EC2のホストCPUにI/Oの負荷をかけない。"
        ),
        (
            "物理ネットワークと\n専用帯域",
            "【EBS最適化帯域 (Dedicated)】\n通常のEC2ネットワーク通信(ENI)とは物理的に隔離された専用リンク(NVMe-oF)でストレージサーバと通信。",
            "【一般ENI帯域 (Shared)】\nバッチ処理がアクセスするRDS(DB)や外部API、S3との通信と同一のネットワーク帯域を奪い合う。",
            "EFSの通信量が増えると、バッチ本来のDBクエリ通信や外部通信の帯域が狭まり、システム全体の遅延を誘発する。"
        ),
        (
            "レプリケーションと\n書き込み合意形成",
            "【同一AZ内の超高速ミラーリング】\n同一データセンター内のペアストレージセルへ書き込み。物理的距離が極小のため 0.5ms で完了。",
            "【3つのAZにまたがる分散合意】\n数十km離れた3つの独立AZ間でクォーラム合意を取ってからACKを返す。物理的な光の速度の限界により数msの遅延が不可避。",
            "ログの追記や一時ファイルという使い捨てのデータに対して、3AZの合意形成を待つのは明らかなアーキテクチャのアンチパターン。"
        )
    ]

    r_arch += 1
    for layer, ebs_mech, efs_mech, insight in arch_data:
        ws3.row_dimensions[r_arch].height = 48
        ws3.cell(row=r_arch, column=1, value=layer).alignment = al_center
        ws3.merge_cells(start_row=r_arch, start_column=2, end_row=r_arch, end_column=3)
        ws3.cell(row=r_arch, column=2, value=ebs_mech).alignment = al_left
        ws3.merge_cells(start_row=r_arch, start_column=4, end_row=r_arch, end_column=5)
        ws3.cell(row=r_arch, column=4, value=efs_mech).alignment = al_left
        ws3.merge_cells(start_row=r_arch, start_column=6, end_row=r_arch, end_column=7)
        ws3.cell(row=r_arch, column=6, value=insight).alignment = al_left

        for c_idx in range(1, 8):
            c = ws3.cell(row=r_arch, column=c_idx)
            c.font = f_normal
            c.fill = fill_zebra_white if r_arch % 2 == 1 else fill_zebra_light
            c.border = border_cell
        r_arch += 1

    # =========================================================================
    # Sheet 4: 04_58JVMバッチ影響度試算
    # =========================================================================
    ws4 = wb.create_sheet(title="04_58JVMバッチ影響度試算")
    ws4.views.sheetView[0].showGridLines = True

    # タイトル
    ws4.merge_cells("A1:G1")
    ws4["A1"].value = "【ワークロード深掘り】本番環境（58JVM並列・32vCPU）における挙動・性能影響シミュレーション"
    ws4["A1"].font = f_title
    ws4["A1"].fill = fill_title
    ws4["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws4.row_dimensions[1].height = 36

    # セクション1: 現状(IT)と本番のリソースプロファイル
    ws4.merge_cells("A3:G3")
    ws4["A3"].value = "1. ワークロードプロファイル分析（IT環境の実測値と本番環境の予測）"
    ws4["A3"].font = f_section
    ws4["A3"].fill = fill_section
    ws4["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws4.row_dimensions[3].height = 24

    profile_headers = ["環境区分", "インスタンス構成", "JVMプロセス数 / 並列度", "CPU使用率と挙動", "ボトルネック分析"]
    ws4.cell(row=4, column=1, value="環境区分").alignment = al_center
    ws4.cell(row=4, column=2, value="インスタンススペック").alignment = al_center
    ws4.cell(row=4, column=3, value="JVM構成 / 並列マルチ数").alignment = al_center
    ws4.merge_cells("D4:E4")
    ws4.cell(row=4, column=4, value="CPU負荷状況とスレッド挙動").alignment = al_center
    ws4.merge_cells("F4:G4")
    ws4.cell(row=4, column=6, value="ワークロード特性の診断結論").alignment = al_center
    ws4.row_dimensions[4].height = 24

    for col in range(1, 8):
        c = ws4.cell(row=4, column=col)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.border = border_th_dark

    profile_data = [
        (
            "現状のIT環境\n(結合テスト)",
            "c7a.2xlarge\n・8vCPU / 16GB RAM\n・Up to 12.5 Gbps 帯域",
            "10 JVM プロセス\n(10マルチ並列実行)",
            "【CPU使用率 90%オーバー】\n1JVMあたり1vCPUをほぼフル消費。8vCPUに対して10JVMが稼働し、CPUスケジューラがフル回転。",
            "典型的な『CPU律速（CPUバウンド）』のバッチ。\n計算処理が極めて重いため、I/O待ちによるスレッドブロックを絶対に発生させてはならない。"
        ),
        (
            "想定される本番環境\n(32vCPU構成)",
            "c7a.8xlarge (想定)\n・32vCPU / 64GB RAM\n・EBS最適化: 12.5 Gbps\n・NW帯域: 12.5 Gbps",
            "58 JVM プロセス\n(32マルチ並列実行)",
            "【CPU使用率 95%〜100%飽和】\n32vCPUに対して58JVMが競合。スレッドコンテキストスイッチが頻発し、CPUリソースは極限状態。",
            "もしディスクI/Oが詰まると、コンテキストスイッチとI/O Waitが二重に爆発し、システム全体がスラッシングに陥る危険性がある。"
        )
    ]

    r_prof = 5
    for env, spec, jvm, cpu, diag in profile_data:
        ws4.row_dimensions[r_prof].height = 50
        ws4.cell(row=r_prof, column=1, value=env).alignment = al_center
        ws4.cell(row=r_prof, column=2, value=spec).alignment = al_left
        ws4.cell(row=r_prof, column=3, value=jvm).alignment = al_center
        ws4.merge_cells(start_row=r_prof, start_column=4, end_row=r_prof, end_column=5)
        ws4.cell(row=r_prof, column=4, value=cpu).alignment = al_left
        ws4.merge_cells(start_row=r_prof, start_column=6, end_row=r_prof, end_column=7)
        ws4.cell(row=r_prof, column=6, value=diag).alignment = al_left

        for c_idx in range(1, 8):
            c = ws4.cell(row=r_prof, column=c_idx)
            c.font = f_normal
            c.fill = fill_zebra_white if r_prof % 2 == 1 else fill_zebra_light
            c.border = border_cell
        r_prof += 1

    # セクション2: ログ＆一時ファイル処理時の挙動シミュレーション
    ws4.merge_cells(f"A{r_prof+1}:G{r_prof+1}")
    ws4[f"A{r_prof+1}"].value = "2. ログ出力および一時ファイル処理時の挙動シミュレーション対比"
    ws4[f"A{r_prof+1}"].font = f_section
    ws4[f"A{r_prof+1}"].fill = fill_section
    ws4[f"A{r_prof+1}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws4.row_dimensions[r_prof+1].height = 24

    sim_headers = ["処理シナリオ", "Java/JVM内部の処理内容", "Amazon EBS (gp3) 採用時の挙動", "Amazon EFS 採用時の挙動", "システムへの影響格差"]
    r_sim = r_prof + 2
    ws4.cell(row=r_sim, column=1, value="シナリオ").alignment = al_center
    ws4.cell(row=r_sim, column=2, value="Java/JVMのI/O処理").alignment = al_center
    ws4.merge_cells(start_row=r_sim, start_column=3, end_row=r_sim, end_column=4)
    ws4.cell(row=r_sim, column=3, value="Amazon EBS (gp3) 採用時").alignment = al_center
    ws4.merge_cells(start_row=r_sim, start_column=5, end_row=r_sim, end_column=6)
    ws4.cell(row=r_sim, column=5, value="Amazon EFS 採用時").alignment = al_center
    ws4.cell(row=r_sim, column=7, value="性能・実行時間の格差").alignment = al_center
    ws4.row_dimensions[r_sim].height = 24

    for col in range(1, 8):
        c = ws4.cell(row=r_sim, column=col)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.border = border_th_dark

    sim_data = [
        (
            "ログ書き出し\n(高頻度Append)",
            "Logback/Log4j2等によるログ追記。\n1行ごとにwrite()およびバッファflush()が発生。\n58JVMで毎秒数万行を出力。",
            "【0.5ms未満で即復帰】\nOSのページキャッシュが即座に受領。EBS専用帯域でバックグラウンド書き込みされるため、Javaスレッドは計算処理を止めずに継続。",
            "【1行ごとに1〜5ms待ち】\nNFS WRITE RPCがネットワークを飛ぶ。58並列のRPCが集中し、JavaのロギングスレッドがBLOCKED状態に陥る。",
            "EFSではログ出力だけでJVM全体の実行速度が数分の一に激突する。"
        ),
        (
            "ログローテーション\n(rename/create)",
            "ファイルサイズまたは時間単位でログファイルをrenameし、新規ファイルを生成(create)。全58JVMで発生。",
            "【数マイクロ秒で完了】\nXFSのディレクトリエントリ更新はメモリ上で即座に完了。他JVMのログ出力に全く影響を与えない。",
            "【NFSディレクトリエントリロック】\nEFS上で全JVMが参照するログディレクトリのメタデータロック競合が発生。数秒間全JVMのログが停止する。",
            "EFSではローテーション時に瞬間的なストップ・ザ・ワールド的フリーズが発生。"
        ),
        (
            "一時ファイル生成\n(Temp File)",
            "File.createTempFile() や\nFiles.createTempFile() による一時ワークファイルの大量作成。",
            "【OSメモリ内でinode即時割当】\nミリ秒未満でファイルディスクリプタが返却され、即座に一時データの書き込み・ソート処理が開始される。",
            "【NFS CREATE RPCの3AZ合意】\n1ファイル作るたびにリモートクラスタとの分散合意が発生(10〜30ms)。58並列で呼ぶと待ち行列が急増。",
            "一時ファイルを多用するバッチでは、EFSは完全にボトルネックとなり処理が停止同等に。"
        ),
        (
            "一時ファイル削除\n(deleteOnExit)",
            "バッチ処理終了時、またはステップ終了時の一時ファイル破棄(unlink/delete)。",
            "【参照カウンタ減算のみ】\nメモリ上のinodeリンクを落とすだけで即完了。ブロック解放はカーネルが非同期に実行。",
            "【NFS REMOVE RPCの分散更新】\n削除操作もメタデータ更新のため3AZ同期待ち。大量のファイル削除処理でプロセス終了が大幅遅延。",
            "バッチ後続処理や次ステップへの移行が、EFSのメタデータ削除待ちで引き延ばされる。"
        ),
        (
            "ネットワーク競合\n(DB・外部API通信)",
            "バッチ処理本体が行うRDSへのSQLクエリ、外部REST API呼び出し、S3とのデータ連携。",
            "【完全無干渉】\nストレージI/Oは「EBS最適化専用帯域」を通るため、一般ENIのネットワーク帯域を1ビットも消費しない。",
            "【帯域の共倒れ】\n大量のログ・一時ファイル通信が一般ENIを圧迫。DBクエリのレスポンスや外部API通信の遅延・タイムアウトを誘発。",
            "EFSを採用すると、ディスクだけでなくDB通信やネットワーク全体が巻き添えを食らう。"
        )
    ]

    r_sim += 1
    for scen, jvm_act, ebs_act, efs_act, gap in sim_data:
        ws4.row_dimensions[r_sim].height = 50
        ws4.cell(row=r_sim, column=1, value=scen).alignment = al_center
        ws4.cell(row=r_sim, column=2, value=jvm_act).alignment = al_left
        ws4.merge_cells(start_row=r_sim, start_column=3, end_row=r_sim, end_column=4)
        ws4.cell(row=r_sim, column=3, value=ebs_act).alignment = al_left
        ws4.merge_cells(start_row=r_sim, start_column=5, end_row=r_sim, end_column=6)
        ws4.cell(row=r_sim, column=5, value=efs_act).alignment = al_left
        ws4.cell(row=r_sim, column=7, value=gap).alignment = al_left

        for c_idx in range(1, 8):
            c = ws4.cell(row=r_sim, column=c_idx)
            c.font = f_normal
            c.fill = fill_zebra_white if r_sim % 2 == 1 else fill_zebra_light
            c.border = border_cell
        r_sim += 1

    # =========================================================================
    # Sheet 5: 05_推奨パラメータ・設計ガイド
    # =========================================================================
    ws5 = wb.create_sheet(title="05_推奨パラメータ・設計ガイド")
    ws5.views.sheetView[0].showGridLines = True

    # タイトル
    ws5.merge_cells("A1:G1")
    ws5["A1"].value = "【実装・運用設計ガイド】本番EC2バッチサーバにおける推奨設定パラメータ"
    ws5["A1"].font = f_title
    ws5["A1"].fill = fill_title
    ws5["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws5.row_dimensions[1].height = 36

    # セクション1: ディスク構成とマウント設計
    ws5.merge_cells("A3:G3")
    ws5["A3"].value = "1. EBSボリューム設計とLinux OSマウントパラメータ（推奨構成）"
    ws5["A3"].font = f_section
    ws5["A3"].fill = fill_section
    ws5["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws5.row_dimensions[3].height = 24

    ws5.cell(row=4, column=1, value="設定対象").alignment = al_center
    ws5.merge_cells("B4:C4")
    ws5.cell(row=4, column=2, value="推奨設定値・パラメータ").alignment = al_center
    ws5.merge_cells("D4:G4")
    ws5.cell(row=4, column=4, value="設定理由とエンジニアリング的効果").alignment = al_center
    ws5.row_dimensions[4].height = 24

    for col in range(1, 8):
        c = ws5.cell(row=4, column=col)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.border = border_th_dark

    mount_params = [
        ("ボリューム構成", "【Multi-EBS構成（ボリューム分離）】\n・Vol 1: ルート(/) 50GB gp3\n・Vol 2: バッチ専用(/data/batch) 500GB+ gp3", "58JVMが一時ファイルやログを大量出力してディスク容量が100%枯渇しても、OS領域が無事であればSSHログインや監視エージェントが死なず、安全に保守・復旧が可能。"),
        ("EBSボリューム種別", "【汎用SSD: gp3】\n・ベースライン: 3,000 IOPS / 125 MB/s (標準)\n・推奨プロビジョン: 6,000 IOPS / 250 MB/s", "gp3は容量と性能を分離可能。バッチ処理の規模に応じて後から無停止でIOPSやスループットを動的に増減できるため、運用柔軟性が極めて高い。"),
        ("ファイルシステム", "【XFS】 (推奨)\nフォーマット: mkfs.xfs -f /dev/nvme1n1", "XFSは並列アロケーション（Allocation Group）アーキテクチャを持ち、58並列のようなマルチプロセスからの同時ファイル割り当て性能がext4より大幅に優れる。"),
        ("マウントオプション\n(/etc/fstab)", "noatime, nodiratime,\nlogbufs=8, logbsize=256k", "【最重要チューニング】\n・noatime: 一時ファイルを読み直すたびに「最終アクセス日時」をディスクに書き換える無駄を100%カット。\n・logbufs/logbsize: XFSジャーナルログバッファを拡大し、大量作成・削除時の競合を抑止。"),
        ("I/Oスケジューラ", "mq-deadline または none\n(cat /sys/block/nvme1n1/queue/scheduler)", "AWS NitroシステムのNVMeデバイスでは、複雑なOSスケジューラを排斥してNitroハードウェアキューに直結させるのが最もレイテンシが低い。")
    ]

    r_mnt = 5
    for tgt, val, eff in mount_params:
        ws5.row_dimensions[r_mnt].height = 44
        ws5.cell(row=r_mnt, column=1, value=tgt).alignment = al_center
        ws5.merge_cells(start_row=r_mnt, start_column=2, end_row=r_mnt, end_column=3)
        ws5.cell(row=r_mnt, column=2, value=val).alignment = al_left
        ws5.merge_cells(start_row=r_mnt, start_column=4, end_row=r_mnt, end_column=7)
        ws5.cell(row=r_mnt, column=4, value=eff).alignment = al_left

        for c_idx in range(1, 8):
            c = ws5.cell(row=r_mnt, column=c_idx)
            c.font = f_normal
            c.fill = fill_zebra_white if r_mnt % 2 == 1 else fill_zebra_light
            c.border = border_cell
        r_mnt += 1

    # セクション2: Java / JVM / アプリケーション設計
    ws5.merge_cells(f"A{r_mnt+1}:G{r_mnt+1}")
    ws5[f"A{r_mnt+1}"].value = "2. Java / JVM / ログ運用設計のベストプラクティス"
    ws5[f"A{r_mnt+1}"].font = f_section
    ws5[f"A{r_mnt+1}"].fill = fill_section
    ws5[f"A{r_mnt+1}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws5.row_dimensions[r_mnt+1].height = 24

    r_jvm = r_mnt + 2
    ws5.cell(row=r_jvm, column=1, value="コンポーネント").alignment = al_center
    ws5.merge_cells("B2:C2")
    ws5.cell(row=r_jvm, column=2, value="設定項目・JVMオプション").alignment = al_center
    ws5.merge_cells("D4:G4")
    ws5.cell(row=r_jvm, column=4, value="推奨設定値・アーキテクチャ詳細").alignment = al_center
    ws5.row_dimensions[r_jvm].height = 24

    for col in range(1, 8):
        c = ws5.cell(row=r_jvm, column=col)
        c.font = f_th_dark
        c.fill = fill_th_dark
        c.border = border_th_dark

    jvm_params = [
        ("一時領域指定\n(java.io.tmpdir)", "-Djava.io.tmpdir=/data/batch/tmp", "デフォルトの /tmp (OSルートボリューム) を避け、バッチ専用EBS(/data/batch/tmp)に逃がす。各JVMプロセスごとにサブディレクトリを分けると競合をさらに防止可能。"),
        ("非同期ロギング\n(Log4j2 / Logback)", "AsyncAppender (DisruptorRingBuffer)\nneverBlock=true, queueSize=262144", "ログ書き出しをJavaのワーカースレッドから完全に切り離し、バックグラウンドのLMAX Disruptorリングバッファで非同期フラッシュ。バッチ処理の計算速度を100%維持。"),
        ("ログ永続化・集約\n(ライフサイクル)", "ローカルEBS出力 ──> CloudWatch Logs / S3\n(Fluent Bit または AWS CLI sync)", "ログはローカルEBSに高速出力し、過去ログはバックグラウンドでAmazon S3やCloudWatch Logsへ非同期アップロード。EFSを使わずに複数サーバログの集中管理を実現。"),
        ("将来のAWS Batch移行\n(アーキテクチャ整合性)", "・一時領域: コンテナローカルEBS利用\n・ログ: STDOUT/STDERR ──> CloudWatch Logs\n・データ連携: Amazon S3 直接連携", "今回「一時領域＝ローカル高速ディスク」「ログ＝非同期集約」としておくことで、将来AWS BatchやECSへコンテナ移行する際、アプリ改修ゼロで自然に移行できる。")
    ]

    r_jvm += 1
    for comp, opt, desc in jvm_params:
        ws5.row_dimensions[r_jvm].height = 44
        ws5.cell(row=r_jvm, column=1, value=comp).alignment = al_center
        ws5.merge_cells(start_row=r_jvm, start_column=2, end_row=r_jvm, end_column=3)
        ws5.cell(row=r_jvm, column=2, value=opt).alignment = al_left
        ws5.merge_cells(start_row=r_jvm, start_column=4, end_row=r_jvm, end_column=7)
        ws5.cell(row=r_jvm, column=4, value=desc).alignment = al_left

        for c_idx in range(1, 8):
            c = ws5.cell(row=r_jvm, column=c_idx)
            c.font = f_normal
            c.fill = fill_zebra_white if r_jvm % 2 == 1 else fill_zebra_light
            c.border = border_cell
        r_jvm += 1

    # =========================================================================
    # 全シートの列幅設定 & ウィンドウ枠固定 (Freeze Panes)
    # =========================================================================
    # Sheet 1 列幅
    col_widths_ws1 = {1: 8, 2: 24, 3: 38, 4: 38, 5: 18, 6: 18, 7: 18}
    for col, width in col_widths_ws1.items():
        ws1.column_dimensions[get_column_letter(col)].width = width
    ws1.freeze_panes = "A5"

    # Sheet 2 列幅
    col_widths_ws2 = {1: 14, 2: 24, 3: 32, 4: 32, 5: 32, 6: 28, 7: 42}
    for col, width in col_widths_ws2.items():
        ws2.column_dimensions[get_column_letter(col)].width = width
    ws2.freeze_panes = "A4"

    # Sheet 3 列幅
    col_widths_ws3 = {1: 16, 2: 32, 3: 32, 4: 20, 5: 20, 6: 22, 7: 22}
    for col, width in col_widths_ws3.items():
        ws3.column_dimensions[get_column_letter(col)].width = width
    ws3.freeze_panes = "A5"

    # Sheet 4 列幅
    col_widths_ws4 = {1: 18, 2: 26, 3: 24, 4: 24, 5: 24, 6: 24, 7: 34}
    for col, width in col_widths_ws4.items():
        ws4.column_dimensions[get_column_letter(col)].width = width
    ws4.freeze_panes = "A5"

    # Sheet 5 列幅
    col_widths_ws5 = {1: 22, 2: 24, 3: 24, 4: 22, 5: 22, 6: 22, 7: 22}
    for col, width in col_widths_ws5.items():
        ws5.column_dimensions[get_column_letter(col)].width = width
    ws5.freeze_panes = "A5"

    # 保存
    output_path = r"C:\Users\taka_\Agy\EFS_EBS_Compare_Agy\EBS_vs_EFS_Batch_Storage_Comparison.xlsx"
    wb.save(output_path)
    print(f"Workbook successfully saved to: {output_path}")

if __name__ == "__main__":
    build_excel_report()
