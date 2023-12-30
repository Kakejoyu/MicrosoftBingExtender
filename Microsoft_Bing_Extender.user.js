// ==UserScript==
// @name        Microsoft Bing Extender
// @name:ja     Microsoft Bing Extender
// @namespace   https://github.com/Kakejoyu/MicrosoftBingExtender
// @version     1.0.1
// @icon        https://www.bing.com/favicon.ico
// @description User script to enhance the functionality of Microsoft Bing
// @description:ja Microsoft Bingの機能を強化するユーザースクリプト
// @author      Kakejoyu
// @supportURL  https://github.com/Kakejoyu/MicrosoftBingExtender/issues
// @match       http*://www.bing.com/*
// @license     MIT
// @grant       GM_registerMenuCommand
// @grant       GM_unregisterMenuCommand
// @grant       GM_setValue
// @grant       GM_getValue
// @grant       GM_addStyle
// @require     https://cdnjs.cloudflare.com/ajax/libs/jquery/3.7.1/jquery.min.js
// @run-at      document-end
// @noframes
// ==/UserScript==

/*! Microsoft Bing Extender | MIT license | https://github.com/Kakejoyu/MicrosoftBingExtender/blob/main/LICENSE */

let lang;

if (document.documentElement.getAttribute('lang') == 'ja') {
  lang = 'ja';
} else {
  lang = 'en';
}

const i18nLib = {
  ja: {
    automateSearch: '自動検索',
    automateSearchOpen: '自動検索を開く',
    gettingFromApiMsg: 'APIから単語を取得中...',
    runningSearchMsg: '自動検索を実行中... [残り：',
    searchFinishedMsg: '自動検索完了！ホームページにリダイレクトします...',
    searchCanceledMsg: '自動検索がキャンセルされました！',
    searchesNumLabel: '検索件数：',
    coolTimeLabel: 'クールタイム',
    coolTimeMinLabel: '最小：',
    coolTimeMaxLabel: '最大：',
    searchStartBtn: '検索開始',
    automateSearchWarning:
      'この機能を使用した場合、MicrosoftはあなたをBANする可能性があります。自己責任で使用してください。私は一切の責任を負いません。',
    automateSearchInfo:
      'ユーザースクリプトでは、ユーザーエージェントを変更することはできません。そのため、モバイル版でポイントを獲得するには、「<a href="https://chromewebstore.google.com/detail/user-agent-switcher-for-c/djflhoibgkdhkhhcedjiklpkjnoahfmg" target="_blank" rel="noopener noreferrer">User-Agent Switcher for Chrome</a>」など、ユーザーエージェントを変更するブラウザ拡張機能をご利用ください。',
    highlightAnswers: '答えをハイライト',
    english: '英語',
    japanese: '日本語',
    darkMode: 'ダークモード',
    lengthLabel: '単語の長さ',
    minLengthLabel: '最小長さ：',
    maxLengthLabel: '最大長さ：',
  },
  en: {
    automateSearch: 'Automate Search',
    automateSearchOpen: 'Open Automate Search',
    gettingFromApiMsg: 'Getting words from API...',
    runningSearchMsg: 'Running Automate search... [Remaining: ',
    searchFinishedMsg: 'Automate search finished! Redirecting to homepage...',
    searchCanceledMsg: 'Automate search canceled!',
    searchesNumLabel: 'Number of Searches:',
    coolTimeLabel: 'Cool Time',
    coolTimeMinLabel: 'Min:',
    coolTimeMaxLabel: 'Max:',
    searchStartBtn: 'Start search',
    automateSearchWarning:
      'Microsoft may ban you for using this feature. Use at your own risk. And I will not take any responsibility.',
    automateSearchInfo:
      'User script cannot change the user agent. Therefore, to earn points on the mobile version, please use a browser extension that changes the user agent, such as "<a href="https://chromewebstore.google.com/detail/user-agent-switcher-for-c/djflhoibgkdhkhhcedjiklpkjnoahfmg" target="_blank" rel="noopener noreferrer">User-Agent Switcher for Chrome</a>".',
    highlightAnswers: 'Highlight Answers',
    english: 'English',
    japanese: 'Japanese',
    darkMode: 'Dark Mode',
    lengthLabel: 'Word length',
    minLengthLabel: 'Min. length:',
    maxLengthLabel: 'Max. length:',
  },
};
const i18n = (key) => i18nLib[lang][key] || `i18n[${lang}][${key}] not found`;

if (GM_getValue('automateSearch', true)) {
  const startAutomateSearch = (
    searchesNum,
    coolTimeMin,
    coolTimeMax,
    minLength,
    maxLength
  ) => {
    GM_setValue('searchesNum', searchesNum);
    GM_setValue('coolTimeMin', coolTimeMin);
    GM_setValue('coolTimeMax', coolTimeMax);
    GM_setValue('minLength', minLength);
    GM_getValue('maxLength', maxLength);
    $('#as-bg').remove();
    $('body').append(
      $(
        `<div id="as-bg">${i18n(
          'gettingFromApiMsg'
        )}<style>#as-bg{position: fixed;top: 5%;left: 25%;width: 50%;padding: 10px;background: #003aa5;color: #fff;border-radius: 10px;z-index: 9999;font-size: 20px;text-align: center;}</style></div>`
      )
    );
    let apiLang;
    if (lang == 'ja') {
      apiLang = 'jp';
    } else {
      apiLang = 'en';
    }
    const request = new XMLHttpRequest();
    request.open(
      'GET',
      `https://random-word.ryanrk.com/api/${apiLang}/word/random/${searchesNum}?minlength=${minLength}&maxlength=${maxLength}`
    );
    request.send();
    request.addEventListener('load', function () {
      const apiData = JSON.parse(this.responseText);
      GM_setValue('words', apiData);
      GM_setValue('isAutomateSearching', true);
      inputSend();
    });
  };

  const inputSend = () => {
    let words = GM_getValue('words');
    if (lang == 'ja') {
      switch (Math.floor(Math.random() * 6)) {
        case 0:
          $('#sb_form_q').val(words[0]);
          break;
        case 1:
          $('#sb_form_q').val(`${words[0]}の意味`);
          break;
        case 2:
          $('#sb_form_q').val(`${words[0]} 意味`);
          break;
        case 3:
          $('#sb_form_q').val(`${words[0]} とは`);
          break;
        case 4:
          $('#sb_form_q').val(`${words[0]} twitter`);
          break;
        case 5:
          $('#sb_form_q').val(`${words[0]} youtube`);
          break;
      }
    } else {
      switch (Math.floor(Math.random() * 5)) {
        case 0:
          $('#sb_form_q').val(words[0]);
          break;
        case 1:
          $('#sb_form_q').val(`Meaning of ${words[0]}`);
          break;
        case 2:
          $('#sb_form_q').val(`What is the meaning of ${words[0]}?`);
          break;
        case 3:
          $('#sb_form_q').val(`What is ${words[0]}?`);
          break;
        case 4:
          $('#sb_form_q').val(`${words[0]} twitter`);
          break;
        case 5:
          $('#sb_form_q').val(`${words[0]} youtube`);
          break;
      }
    }
    words.shift();
    if (words.length == 0) {
      GM_setValue('isAutomateSearching', false);
      GM_setValue('isFinished', true);
    }
    GM_setValue('words', words);
    $('#sb_form').submit();
  };

  let searchTimer;
  const automateSearch = () => {
    const coolTimeMin = GM_getValue('coolTimeMin') * 1000;
    const coolTimeMax = GM_getValue('coolTimeMax') * 1000;
    if (GM_getValue('isAutomateSearching')) {
      $('#as-bg').remove();
      $('body').append(
        $(
          `<div id="as-bg">${
            i18n('runningSearchMsg') +
            GM_getValue('words').length +
            ' / ' +
            GM_getValue('searchesNum') +
            ']'
          }<svg xmlns="http://www.w3.org/2000/svg" height="1em" viewBox="0 0 512 512" id="as-cancel-btn"><!--! Font Awesome Free 6.4.2 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. --><path d="M464 256A208 208 0 1 0 48 256a208 208 0 1 0 416 0zM0 256a256 256 0 1 1 512 0A256 256 0 1 1 0 256zm192-96H320c17.7 0 32 14.3 32 32V320c0 17.7-14.3 32-32 32H192c-17.7 0-32-14.3-32-32V192c0-17.7 14.3-32 32-32z"/></svg>
                    <style>#as-bg{position: fixed;top: 5%;left: 25%;width: 50%;padding: 10px;background: #003aa5;color: #fff;border-radius: 10px;z-index: 9999;font-size: 20px;text-align: center;}#as-cancel-btn{fill: #ffffff;margin-left: 10px;font-size: 26px;vertical-align: -5px;cursor: pointer;}</style></div>`
        )
      );
      $('#as-cancel-btn').click(() => {
        clearTimeout(searchTimer);
        GM_setValue('isAutomateSearching', false);
        GM_setValue('isFinished', false);
        GM_setValue('words', []);
        $('#as-bg').remove();
        $('body').append(
          $(`<div id="as-bg">${i18n('searchCanceledMsg')}
                <svg xmlns="http://www.w3.org/2000/svg" height="1em" viewBox="0 0 512 512" id="as-close-btn">
                    <!--! Font Awesome Free 6.4.2 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. -->
                    <path d="M256 48a208 208 0 1 1 0 416 208 208 0 1 1 0-416zm0 464A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM175 175c-9.4 9.4-9.4 24.6 0 33.9l47 47-47 47c-9.4 9.4-9.4 24.6 0 33.9s24.6 9.4 33.9 0l47-47 47 47c9.4 9.4 24.6 9.4 33.9 0s9.4-24.6 0-33.9l-47-47 47-47c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-47 47-47-47c-9.4-9.4-24.6-9.4-33.9 0z"></path>
                </svg>
                <style>
                    #as-bg{position: fixed;top: 5%;left: 25%;width: 50%;padding: 10px;background: rgb(196, 0, 0);color: #fff;border-radius: 10px;z-index: 9999;font-size: 20px;text-align: center;}
                    #as-close-btn{fill: #ffffff;margin-left: 10px;font-size: 26px;vertical-align: -5px;cursor: pointer;}
                </style>
            </div>`)
        );
        $('#as-close-btn').click(() => {
          $('#as-bg').remove();
        });
      });
      searchTimer = setTimeout(
        inputSend,
        Math.floor(Math.random() * (coolTimeMax - coolTimeMin)) + coolTimeMin
      );
    } else {
      if (GM_getValue('isFinished')) {
        GM_setValue('isFinished', false);
        $('#as-bg').remove();
        $('body').append(
          $(
            `<div id="as-bg">${i18n(
              'searchFinishedMsg'
            )}<style>#as-bg{position: fixed;top: 5%;left: 25%;width: 50%;padding: 10px;background: #00a500;color: #fff;border-radius: 10px;z-index: 9999;font-size: 20px;text-align: center;}</style></div>`
          )
        );
        setTimeout(() => {
          location.href = 'https://www.bing.com';
        }, Math.floor(Math.random() * (coolTimeMax - coolTimeMin)) + coolTimeMin);
      }
    }
  };

  automateSearch();

  const showAutomateSearchUI = () => {
    let UIcss;
    if (GM_getValue('darkMode', true)) {
      UIcss = `body {overflow: hidden;}
#as-bg {position: fixed;z-index: 999999;background-color: rgba(0, 0, 0, 0.8);left: 0px;top: 0px;user-select: none;-moz-user-select: none;}
#as-fg {width: 50%;height: 82%;padding: 15px;position: absolute;top: 10%;left: 25%;background: #111111;border-radius: 20px;}
#as-fg * {margin: 7px 0;font-family: sans-serif;font-size: 15px;color: #ffffff;}
#as-close {position: absolute;right: 10px;top: 10px;width: 32px;height: 32px;cursor: pointer;fill: currentColor;}
#as-fg p.as-warning{background: rgb(196, 0, 0);color: #fff;padding: 10px;border-radius: 5px;}
#as-fg p.as-info{background: rgb(0 86 196);color: #fff;padding: 10px;border-radius: 5px;}
#as-fg p.as-info a{text-decoration: underline;font-weight: bold;}
#as-fg label{display: inline-block;}
#as-fg .as-title{font-size: 30px;font-weight: bold;}
#as-fg .as-bold{font-weight: bold;}
#as-fg .as-input{background: #111111;border: 3px solid #666;border-radius: 5px;height: 30px;padding: 5px;font-size: 25px;width: 60px;}
#as-fg .as-input:hover{background: #2b2b2b;}
#as-fg .as-input:focus{background: #3a3a3a;}
#as-fg #as-search-start-btn{background: #1b1b1b;border: 3px solid #666;border-radius: 5px;height: 45px;width: 100%;padding: 5px;font-size: 25px;cursor: pointer;}
#as-fg #as-search-start-btn:hover{background: #2b2b2b;}`;
    } else {
      UIcss = `body {overflow: hidden;}
#as-bg {position: fixed;z-index: 999999;background-color: rgba(0, 0, 0, 0.8);left: 0px;top: 0px;user-select: none;-moz-user-select: none;}
#as-fg {width: 50%;height: 82%;padding: 15px;position: absolute;top: 10%;left: 25%;background: #ffffff;border-radius: 20px;}
#as-fg * {margin: 7px 0;font-family: sans-serif;font-size: 15px;color: #111111;}
#as-close {position: absolute;right: 10px;top: 10px;width: 32px;height: 32px;cursor: pointer;fill: currentColor;}
#as-fg p.as-warning{background: rgb(196, 0, 0);color: #fff;padding: 10px;border-radius: 5px;}
#as-fg p.as-info{background: rgb(0 86 196);color: #fff;padding: 10px;border-radius: 5px;}
#as-fg p.as-info a{color: #fff;text-decoration: underline;font-weight: bold;}
#as-fg label{display: inline-block;}
#as-fg .as-title{font-size: 30px;font-weight: bold;}
#as-fg .as-bold{font-weight: bold;}
#as-fg .as-input{background: #ffffff;border: 3px solid #666;border-radius: 5px;height: 30px;padding: 5px;font-size: 25px;width: 60px;}
#as-fg .as-input:hover{background: #dddddd;}
#as-fg .as-input:focus{background: #b3b3b3;}
#as-fg #as-search-start-btn{background: #cacaca;border: 3px solid #666;border-radius: 5px;height: 45px;width: 100%;padding: 5px;font-size: 25px;cursor: pointer;}
#as-fg #as-search-start-btn:hover{background: #a1a1a1;}`;
    }
    $('#as-bg').remove();
    let bg = $(`<div id="as-bg">
    <div id="as-fg">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" id="as-close">
            <!--! Font Awesome Free 6.4.2 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. -->
            <path d="M256 48a208 208 0 1 1 0 416 208 208 0 1 1 0-416zm0 464A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM175 175c-9.4 9.4-9.4 24.6 0 33.9l47 47-47 47c-9.4 9.4-9.4 24.6 0 33.9s24.6 9.4 33.9 0l47-47 47 47c9.4 9.4 24.6 9.4 33.9 0s9.4-24.6 0-33.9l-47-47 47-47c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-47 47-47-47c-9.4-9.4-24.6-9.4-33.9 0z"></path>
        </svg>
        <p class="as-title">${i18n(
          'automateSearch'
        )}</p><p class="as-warning">${i18n('automateSearchWarning')}</p>
        <p class="as-info">${i18n('automateSearchInfo')}</p>
        <label>${i18n(
          'searchesNumLabel'
        )}<input class="as-input" type="number" value="${GM_getValue(
      'searchesNum',
      '60'
    )}" step="1" min="1" id="as-searches-num"></label><br><span class="as-bold">${i18n(
      'coolTimeLabel'
    )}</span><br>
        <label>${i18n(
          'coolTimeMinLabel'
        )}<input class="as-input" type="number" value="${GM_getValue(
      'coolTimeMin',
      '4'
    )}" min="1" id="as-cool-time-min"></label><label>${i18n(
      'coolTimeMaxLabel'
    )}<input class="as-input" type="number" value="${GM_getValue(
      'coolTimeMax',
      '7'
    )}" min="1" id="as-cool-time-max"></label><br>
        <span class="as-bold">${i18n('lengthLabel')}</span><br>
        <label>${i18n(
          'maxLengthLabel'
        )}<input class="as-input" type="number" value="${GM_getValue(
      'minLength',
      '4'
    )}" min="1" id="as-min-length"></label><label>${i18n(
      'minLengthLabel'
    )}<input class="as-input" type="number" value="${GM_getValue(
      'maxLength',
      '7'
    )}" min="1" id="as-max-length"></label>
        <button id="as-search-start-btn">${i18n('searchStartBtn')}</button>
    </div>
    <style>${UIcss}</style>
</div>`);
    $('body').append(bg);
    $('#as-bg').css({
      width: document.documentElement.clientWidth + 'px',
      height: document.documentElement.clientHeight + 'px',
    });
    $('#as-close').click(function () {
      $('#as-bg').remove();
    });
    $('#as-search-start-btn').click(() => {
      startAutomateSearch(
        $('#as-searches-num').val(),
        $('#as-cool-time-min').val(),
        $('#as-cool-time-max').val(),
        $('#as-min-length').val(),
        $('#as-max-length').val()
      );
    });
  };
  GM_registerMenuCommand(i18n('automateSearchOpen'), () => {
    showAutomateSearchUI();
  });
}

let menuId = [];
const registerMenu = () => {
  if (menuId.length) {
    const len = menuId.length;
    for (let i = 0; i < len; i++) {
      GM_unregisterMenuCommand(menuId[i]);
    }
  }
  const menu = [
    ['automateSearch', true],
    ['highlightAnswers', true],
    ['darkMode', true],
  ];
  const len = menu.length;
  for (let i = 0; i < len; i++) {
    const item = menu[i][0];
    menu[i][1] = GM_getValue(item);
    if (menu[i][1] === null || menu[i][1] === undefined) {
      GM_setValue(item, true);
      menu[i][1] = true;
    }
    menuId[i] = GM_registerMenuCommand(
      `${menu[i][1] ? '✅' : '❌'} ${i18n(item)}`,
      () => {
        GM_setValue(item, !menu[i][1]);
        registerMenu();
      }
    );
  }

  return Object.freeze({
    automateSearch: menu[0][1],
    highlightAnswers: menu[1][1],
    darkMode: menu[2][1],
  });
};
const config = registerMenu();

if (config.highlightAnswers) {
  if (config.darkMode) {
    GM_addStyle(`/* [Microsoft Rewards] Color in the correct answer */
div#overlayWrapper[aria-label*="Microsoft Rewards"] div[iscorrectoption="True"] {
    background-color: #243b28 !important;
}

/* [Microsoft Rewards] Color in the incorrect answers */
div#overlayWrapper[aria-label*="Microsoft Rewards"] div[iscorrectoption="False"] {
    background-color: #853a3a !important;
}`);
  } else {
    GM_addStyle(`/* [Microsoft Rewards] Color in the correct answer */
div#overlayWrapper[aria-label*="Microsoft Rewards"] div[iscorrectoption="True"] {
    background-color: #75f08a !important;
}

/* [Microsoft Rewards] Color in the incorrect answers */
div#overlayWrapper[aria-label*="Microsoft Rewards"] div[iscorrectoption="False"] {
    background-color: #f07575 !important;
}`);
  }
}
