NDLOCR-Lite のモデル（ブラウザ用）

出典：国立国会図書館 NDLラボ「NDLOCR-Lite」 https://github.com/ndl-lab/ndlocr-lite
ライセンス：CC BY 4.0（https://creativecommons.org/licenses/by/4.0/deed.ja）

  deim.onnx       ← src/model/deim-s-1024x1024.onnx（レイアウト・行の検出）
  parseq30.onnx   ← src/model/parseq-ndl-24x256-30-tiny-189epoch-tegaki3-r8data-202604.onnx
  parseq50.onnx   ← src/model/parseq-ndl-24x384-50-tiny-300epoch-tegaki3-r8data-202604.onnx
  parseq100.onnx  ← src/model/parseq-ndl-24x768-100-tiny-153epoch-tegaki3-r8data-202604.onnx
  chars.json      ← src/config/NDLmoji.yaml の文字一覧

変更点：容量を減らすため、重み（1024要素以上の浮動小数点の定数）を16bitで保存し、
読み込み時に32bitへ戻す Cast ノードを加えました。計算は元のモデルと同じ32bitで行います。
（同梱の写真4枚で、元のモデルと読み取り結果が一致することを確認済み）

../ort/ は ONNX Runtime Web 1.26.0（MIT License, Microsoft）の配布ファイルです。
