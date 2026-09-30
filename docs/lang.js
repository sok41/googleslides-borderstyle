// 表示言語：ブラウザの言語が日本語なら日本語、それ以外は英語。ヘッダーのボタンで切り替えられる
(function () {
  const root = document.documentElement;
  const initial = /^ja/i.test(navigator.language || '') ? 'ja' : 'en';

  function setLang(lang) {
    root.classList.remove('lang-ja', 'lang-en');
    root.classList.add('lang-' + lang);
    root.lang = lang;
    document.querySelectorAll('.lang-switch button').forEach(b => {
      b.setAttribute('aria-pressed', String(b.dataset.lang === lang));
    });
  }

  document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.lang-switch button').forEach(b => {
      b.addEventListener('click', () => setLang(b.dataset.lang));
    });
    setLang(initial);
  });
  setLang(initial);
})();
