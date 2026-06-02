import { _buildPrompt, _parseGeminiResponse, GeminiResponse } from '../ai-review-analyzer';

// ── _buildPrompt ─────────────────────────────────────────────

describe('_buildPrompt', () => {
  test('レビューテキストがプロンプトに含まれる', () => {
    const result = _buildPrompt('商品の品質が良かったです');
    expect(result).toContain('商品の品質が良かったです');
  });

  test('JSON出力形式の指示が含まれる', () => {
    const result = _buildPrompt('テストレビュー');
    expect(result).toContain('"sentiment"');
    expect(result).toContain('"summary"');
  });

  test('有効な感情分類ラベルが列挙されている', () => {
    const result = _buildPrompt('テストレビュー');
    expect(result).toContain('ポジティブ');
    expect(result).toContain('ネガティブ');
    expect(result).toContain('中立');
  });

  test('空文字を渡してもプロンプト文字列が返る', () => {
    const result = _buildPrompt('');
    expect(typeof result).toBe('string');
    expect(result.length).toBeGreaterThan(0);
  });
});

// ── _parseGeminiResponse ─────────────────────────────────────

function makeResponse(text: string): GeminiResponse {
  return {
    candidates: [{ content: { parts: [{ text }] } }],
  };
}

describe('_parseGeminiResponse', () => {
  test('正常なレスポンスを正しくパースする（ポジティブ）', () => {
    const json = makeResponse('{"sentiment": "ポジティブ", "summary": "良い商品でした"}');
    const result = _parseGeminiResponse(json);
    expect(result).not.toBeNull();
    expect(result?.sentiment).toBe('ポジティブ');
    expect(result?.summary).toBe('良い商品でした');
  });

  test('正常なレスポンスを正しくパースする（ネガティブ）', () => {
    const json = makeResponse('{"sentiment": "ネガティブ", "summary": "品質が悪かった"}');
    const result = _parseGeminiResponse(json);
    expect(result?.sentiment).toBe('ネガティブ');
  });

  test('正常なレスポンスを正しくパースする（中立）', () => {
    const json = makeResponse('{"sentiment": "中立", "summary": "普通でした"}');
    const result = _parseGeminiResponse(json);
    expect(result?.sentiment).toBe('中立');
  });

  test('JSON以外のテキストが前後にあっても抽出できる', () => {
    const json = makeResponse('以下が結果です\n{"sentiment": "ポジティブ", "summary": "良い"}\n以上です');
    const result = _parseGeminiResponse(json);
    expect(result).not.toBeNull();
    expect(result?.sentiment).toBe('ポジティブ');
  });

  test('JSONが含まれない場合は null を返す', () => {
    const json = makeResponse('JSONではないテキストのみ');
    const result = _parseGeminiResponse(json);
    expect(result).toBeNull();
  });

  test('無効な感情分類値の場合は null を返す', () => {
    const json = makeResponse('{"sentiment": "unknown", "summary": "テスト"}');
    const result = _parseGeminiResponse(json);
    expect(result).toBeNull();
  });

  test('summary が省略された場合は空文字列になる', () => {
    const json = makeResponse('{"sentiment": "中立"}');
    const result = _parseGeminiResponse(json);
    expect(result).not.toBeNull();
    expect(result?.summary).toBe('');
  });

  test('不正なJSONの場合は null を返す', () => {
    const json = makeResponse('{invalid json}');
    const result = _parseGeminiResponse(json);
    expect(result).toBeNull();
  });
});
