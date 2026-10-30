document.getElementById('year').textContent = new Date().getFullYear();

// 吹き出しの「詳しく見る」は ?section=<節のid> で開く。いきなりその節を見せず、
// ページの先頭を一瞬見せてから滑らかにスクロールする。上にも節があることに気づけるようにするため。
// # で渡さないのは、ブラウザが読み込みの最後に自分でその節へ飛び、先頭から動かせないため。
// 視差効果を減らす設定（prefers-reduced-motion）の人には動かさず、その節へ直接移る。
(() => {
  const id = new URLSearchParams(location.search).get('section');
  const target = id && document.getElementById(id);
  if (!target) return;

  if (matchMedia('(prefers-reduced-motion: reduce)').matches) {
    target.scrollIntoView({ block: 'start' });
    return;
  }
  window.addEventListener('load', () => {
    setTimeout(() => target.scrollIntoView({ behavior: 'smooth', block: 'start' }), 600);
  });
})();
