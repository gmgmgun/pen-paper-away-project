const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

// 폼(ppap_form.html)은 DOM 하네스가 없어 소스에서 초기값 등록 분기의 비활성화 카드 목록을 읽는다.
// 초기값 등록은 계기판 수치가 필수다(제출 시 "계기판 수치를 입력해주세요.", 결과 화면 "초기 계기판").
function firstRecordDisabledCards() {
  const src = fs.readFileSync(path.join(__dirname, "..", "ppap_form.html"), "utf8");
  const start = src.indexOf("state.isFirst = true;");
  assert.notEqual(start, -1, "초기값 등록 분기를 찾지 못함");
  const list = src.slice(start).match(/\[([^\]]*)\]\.forEach\(/);
  assert.ok(list, "비활성화 카드 목록을 찾지 못함");
  return [...list[1].matchAll(/"([^"]+)"/g)].map((m) => m[1]);
}

test("초기값 등록 모드는 계기판 카드를 비활성화하지 않는다", () => {
  assert.ok(!firstRecordDisabledCards().includes("fieldOdoCard"));
});

test("초기값 등록 모드는 성명·사용구분·비고 카드만 비활성화한다", () => {
  assert.deepEqual(firstRecordDisabledCards(), [
    "fieldNameCard",
    "fieldTypeCard",
    "fieldNoteCard",
  ]);
});
