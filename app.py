import os
import json
import math
import base64
import re
from flask import Flask, render_template, request, jsonify
import anthropic

app = Flask(__name__)
client = anthropic.Anthropic()


def heron(a, b, c):
    s = (a + b + c) / 2
    val = s * (s - a) * (s - b) * (s - c)
    return math.sqrt(val) if val >= 0 else None


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    if "file" not in request.files or not request.files["file"].filename:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    file = request.files["file"]
    tolerance = float(request.form.get("tolerance", "0.005"))

    file_bytes = file.read()
    b64 = base64.standard_b64encode(file_bytes).decode("utf-8")

    fname = file.filename.lower()
    if fname.endswith(".pdf"):
        media_type = "application/pdf"
        block_type = "document"
    elif fname.endswith(".png"):
        media_type = "image/png"
        block_type = "image"
    elif fname.endswith(".webp"):
        media_type = "image/webp"
        block_type = "image"
    else:
        media_type = "image/jpeg"
        block_type = "image"

    prompt = """この図面・計算表は三斜求積図です。

【読み取り指示】
三斜求積法では、複雑な形状の土地を複数の三角形に分割し、各三角形の三辺(a, b, c)からヘロンの公式で面積を求めます。

以下を全て読み取ってください：
1. 各三角形の番号（No.）
2. 各三角形の三辺の長さ（a, b, c）※単位: m
3. 計算表に業者が記載した各三角形の面積（単位: m²）
4. 合計面積（記載がある場合）

【出力形式】
以下のJSON形式のみで返答（説明文・マークダウン不要）：

{
  "triangles": [
    {"no": 1, "a": 10.500, "b": 8.200, "c": 12.100, "stated_area": 42.123},
    {"no": 2, "a": 5.000, "b": 6.000, "c": 7.000, "stated_area": 14.697}
  ],
  "stated_total": 56.820
}

【注意】
- 読み取れない値はnullとする
- 三辺が特定できない場合、読み取れた辺をa, b, cに割り当てる
- 合計面積が記載されていない場合、stated_totalはnull"""

    if block_type == "document":
        file_block = {
            "type": "document",
            "source": {"type": "base64", "media_type": media_type, "data": b64},
        }
    else:
        file_block = {
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": b64},
        }

    try:
        response = client.messages.create(
            model="claude-opus-4-7",
            max_tokens=4096,
            messages=[
                {
                    "role": "user",
                    "content": [file_block, {"type": "text", "text": prompt}],
                }
            ],
        )
    except anthropic.APIError as e:
        return jsonify({"error": f"APIエラー: {str(e)}"}), 500

    raw = response.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw).strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return jsonify({"error": "図面の読み取り結果を解析できませんでした", "raw": raw}), 500

    results = []
    calc_total = 0.0

    for tri in data.get("triangles", []):
        a, b, c, stated = (
            tri.get("a"),
            tri.get("b"),
            tri.get("c"),
            tri.get("stated_area"),
        )
        calc_area = diff = ok = None
        err = None

        if all(v is not None for v in (a, b, c)):
            try:
                a, b, c = float(a), float(b), float(c)
                if a + b > c and a + c > b and b + c > a:
                    calc_area = heron(a, b, c)
                    if calc_area is not None:
                        calc_total += calc_area
                        if stated is not None:
                            stated = float(stated)
                            diff = calc_area - stated
                            ok = abs(diff) <= tolerance
                else:
                    err = "三角不等式不成立"
            except (ValueError, TypeError):
                err = "数値変換エラー"

        results.append(
            {
                "no": tri.get("no"),
                "a": a,
                "b": b,
                "c": c,
                "stated_area": float(stated) if stated is not None else None,
                "calc_area": round(calc_area, 4) if calc_area is not None else None,
                "diff": round(diff, 4) if diff is not None else None,
                "ok": ok,
                "error": err,
            }
        )

    stated_total = data.get("stated_total")
    total_ok = total_diff = None
    if stated_total is not None:
        stated_total = float(stated_total)
        total_diff = calc_total - stated_total
        total_ok = abs(total_diff) <= tolerance

    return jsonify(
        {
            "triangles": results,
            "calc_total": round(calc_total, 4),
            "stated_total": stated_total,
            "total_diff": round(total_diff, 4) if total_diff is not None else None,
            "total_ok": total_ok,
            "tolerance": tolerance,
        }
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
