document.getElementById('year').textContent = new Date().getFullYear();

// 吹き出しの「詳しく見る」は ?section=<節のid> で開く。いきなりその節を見せず、
// ページの先頭を一瞬見せてから滑らかにスクロールする。上にも節があることに気づけるようにするため。
// # で渡さないのは、ブラウザが読み込みの最後に自分でその節へ飛び、先頭から動かせないため。
// 視差効果を減らす設定（prefers-reduced-motion）の人には動かさず、その節へ直接移る。
(() => {
  const id = new URLSearchParams(location.search).get('section');
  const target = id && document.getElementById(id);
  if (!target) return;

  // どちらも画像の読み込みが終わってから動かす。先に動かすと、上の画像の高さが後から伸びて着地点がずれる
  const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;
  window.addEventListener('load', () => {
    if (reduced) {
      target.scrollIntoView({ block: 'start' });
      return;
    }
    setTimeout(() => target.scrollIntoView({ behavior: 'smooth', block: 'start' }), 600);
  });
})();
