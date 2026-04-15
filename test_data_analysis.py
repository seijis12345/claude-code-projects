"""
data-analysis環境の動作確認スクリプト
"""
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # GUIなしで動作
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# Windows日本語フォント設定（Meiryo / Yu Gothic）
_jp_fonts = [f.name for f in fm.fontManager.ttflist if any(
    k in f.name for k in ('Meiryo', 'Yu Gothic', 'MS Gothic', 'IPAexGothic')
)]
if _jp_fonts:
    plt.rcParams['font.family'] = _jp_fonts[0]

print("=== ライブラリバージョン確認 ===")
print(f"numpy:      {np.__version__}")
print(f"pandas:     {pd.__version__}")
print(f"matplotlib: {matplotlib.__version__}")

# --- numpy ---
print("\n=== numpy テスト ===")
arr = np.array([1, 2, 3, 4, 5])
print(f"配列: {arr}")
print(f"平均: {arr.mean()}, 標準偏差: {arr.std():.4f}")

# --- pandas ---
print("\n=== pandas テスト ===")
df = pd.DataFrame({
    "名前":   ["Alice", "Bob", "Carol"],
    "点数":   [85, 92, 78],
    "合否":   ["合格", "合格", "合格"],
})
print(df)
print(f"\n平均点: {df['点数'].mean()}")

# --- matplotlib ---
print("\n=== matplotlib テスト ===")
fig, ax = plt.subplots()
ax.plot([1, 2, 3, 4], [10, 20, 15, 25], marker='o', label="サンプルデータ")
ax.set_title("動作確認グラフ")
ax.set_xlabel("X軸")
ax.set_ylabel("Y軸")
ax.legend()
output_path = "C:/Users/seijis/Claude_Code/test_plot.png"
fig.savefig(output_path)
plt.close()
print(f"グラフを保存しました: {output_path}")

print("\n=== 全テスト完了 ===")
