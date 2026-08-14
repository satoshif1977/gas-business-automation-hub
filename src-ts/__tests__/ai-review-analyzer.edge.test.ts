import { _buildPrompt, _parseGeminiResponse, GeminiResponse } from '../ai-review-analyzer';

// ─── ヘルパー ─────────────────────────────────────────────

function makeResponse(text: string): GeminiResponse {
  return {
    candidates: [{ content: { parts: [{ text }] } }],
  };
}

// ── _buildPrompt エッジケース ─────────────────────────────

describe('_buildPrompt edge cases', () => {
  test('プロンプトに "顧客レビュー" が含まれる', () => {
    const result = _buildPrompt('テスト');
    expect(result).toContain('顧客レビュー');
  });

  test('プロンプトに "30文字以内" が含まれる', () => {
    const result = _buildPrompt('テスト');
    expect(result).toContain('30文字以内');
  });

  test('改行を含むレビューテキストがそのまま埋め込まれる', () => {
    const reviewWithNewline = '品質が良かった\nまた購入したい';
    const result = _buildPrompt(reviewWithNewline);
    expect(result).toContain('品質が良かった\nまた購入したい');
  });

  test('特殊文字（「」・）を含むレビューも埋め込まれる', () => {
    const review = '「大変」良い商品・サービスでした';
    const result = _buildPrompt(review);
    expect(result).toContain(review);
  });

  test('プロンプトは文字列型で 50文字以上の長さを持つ', () => {
    const result = _buildPrompt('テスト');
    expect(typeof result).toBe('string');
    expect(result.length).toBeGreaterThan(50);
  });

  test('長いレビューテキスト（500文字）もプロンプトに含まれる', () => {
    const longReview = 'あ'.repeat(500);
    const result = _buildPrompt(longReview);
    expect(result).toContain(longReview);
  });
});

// ── _parseGeminiResponse エッジケース ─────────────────────

describe('_parseGeminiResponse edge cases', () => {
  test('JSON の前後に余分な改行があっても正常にパースできる', () => {
    const result = _parseGeminiResponse(
      makeResponse('\n\n{"sentiment": "中立", "summary": "普通"}\n\n')
    );
    expect(result?.sentiment).toBe('中立');
    expect(result?.summary).toBe('普通');
  });

  test('sentiment と summary 以外のフィールドがあっても正常にパースできる', () => {
    const result = _parseGeminiResponse(
      makeResponse('{"sentiment": "ポジティブ", "summary": "良い", "extra": "ignored"}')
    );
    expect(result?.sentiment).toBe('ポジティブ');
  });

  test('text が空文字の場合は null を返す', () => {
    const result = _parseGeminiResponse(makeResponse(''));
    expect(result).toBeNull();
  });

  test('英語の sentiment は無効として null を返す', () => {
    const result = _parseGeminiResponse(
      makeResponse('{"sentiment": "positive", "summary": "good"}')
    );
    expect(result).toBeNull();
  });

  test('sentiment が null の場合は null を返す', () => {
    const result = _parseGeminiResponse(
      makeResponse('{"sentiment": null, "summary": "テスト"}')
    );
    expect(result).toBeNull();
  });

  test('summary が空文字の場合はそのまま返す', () => {
    const result = _parseGeminiResponse(
      makeResponse('{"sentiment": "ネガティブ", "summary": ""}')
    );
    expect(result?.summary).toBe('');
  });

  test('candidates が空配列の場合は null を返す', () => {
    const emptyResp: GeminiResponse = { candidates: [] };
    expect(_parseGeminiResponse(emptyResp)).toBeNull();
  });

  test('全3種の sentiment がそれぞれ正しく返る', () => {
    const labels = ['ポジティブ', 'ネガティブ', '中立'] as const;
    for (const label of labels) {
      const result = _parseGeminiResponse(
        makeResponse(`{"sentiment": "${label}", "summary": "テスト"}`)
      );
      expect(result?.sentiment).toBe(label);
    }
  });
});
