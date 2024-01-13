// ==UserScript==
// @name        Microsoft Bing Extender
// @name:ja     Microsoft Bing Extender
// @namespace   https://github.com/Kakejoyu/MicrosoftBingExtender
// @version     1.1.0
// @icon        https://www.bing.com/favicon.ico
// @description     User script to enhance the functionality of Microsoft Bing
// @description:ja  Microsoft Bingの機能を強化するユーザースクリプト
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

jQuery(($) => {
  'use strict';

  let lang;

  if (document.documentElement.getAttribute('lang') === 'ja') {
    lang = 'ja';
  } else {
    lang = 'en';
  }

  const i18nLib = {
    ja: {
      automateSearch: '自動検索',
      highlightAnswers: '答えをハイライト',
      darkMode: 'ダークモード',
      reload: 'リロードして設定を適用',
      automateSearchOpen: '自動検索を開く',
      gettingFromApiMsg: 'APIから単語を取得中...',
      runningMsg:
        '実行中... [ユニット %0 / %1：残り %2 / %3] [合計：残り %4 / %5]',
      coolDownMsg:
        'クールダウン中... [次のユニット %0 / %1 まであと <span id="mbe-countdown"></span>] [合計：残り %2 / %3]',
      searchFinishedMsg: '完了！ [計 %0 ユニット] [計 %1 回検索]',
      searchCanceledMsg: '自動検索がキャンセルされました！',
      lengthLabel: '単語の長さ (文字数)',
      minLengthLabel: '最小：',
      maxLengthLabel: '最大：',
      iuIntervalLabel: 'ユニット内検索間隔 (秒)',
      iuIntervalMinLabel: '最小：',
      iuIntervalMaxLabel: '最大：',
      searchesNumPULabel: '1ユニットあたりの検索件数：',
      unitInterval: 'ユニット間隔 (分)：',
      unitCount: 'ユニット数：',
      totalSearchesLabel: '検索数合計：%0 回',
      estimatedTimeLabel: '予想所要時間：',
      searchStartBtn: '検索開始',
      automateSearchWarning:
        'この機能を使用した場合、MicrosoftはあなたをBANする可能性があります。自己責任で使用してください。私は一切の責任を負いません。',
      automateSearchInfo:
        'ユーザースクリプトでは、ユーザーエージェントを変更することはできません。そのため、モバイル版でポイントを獲得するには、「<a href="https://chromewebstore.google.com/detail/user-agent-switcher-for-c/djflhoibgkdhkhhcedjiklpkjnoahfmg" target="_blank" rel="noopener noreferrer">User-Agent Switcher for Chrome</a>」など、ユーザーエージェントを変更するブラウザ拡張機能をご利用ください。',
    },
    en: {
      automateSearch: 'Automate Search',
      highlightAnswers: 'Highlight Answers',
      darkMode: 'Dark Mode',
      reload: 'Reload and apply settings',
      automateSearchOpen: 'Open Automate Search',
      gettingFromApiMsg: 'Getting words from API...',
      runningMsg:
        'Running... [Unit %0 / %1: Remaining %2 / %3] [Total: Remaining %4 / %5]',
      coolDownMsg:
        'On cool down... [<span id="mbe-countdown"></span> remaining until next %0 / %1 unit ] [Total: %2 / %3 remaining]',
      searchFinishedMsg: 'Completed! [total %0 units] [total %1 searches]',
      searchCanceledMsg: 'Automate search canceled!',
      lengthLabel: 'Word length (character count)',
      minLengthLabel: 'Min: ',
      maxLengthLabel: 'Max: ',
      iuIntervalLabel: 'Intra-unit search interval (seconds)',
      iuIntervalMinLabel: 'Min: ',
      iuIntervalMaxLabel: 'Max: ',
      searchesNumPULabel: 'Number of searches per unit: ',
      unitInterval: 'Unit interval (minutes): ',
      unitCount: 'Unit Count: ',
      totalSearchesLabel: 'Total number of searches: %0',
      estimatedTimeLabel: 'Estimated time required: ',
      searchStartBtn: 'Start search',
      automateSearchWarning:
        'Microsoft may ban you for using this feature. Use at your own risk. And I will not take any responsibility.',
      automateSearchInfo:
        'User script cannot change the user agent. Therefore, to earn points on the mobile version, please use a browser extension that changes the user agent, such as "<a href="https://chromewebstore.google.com/detail/user-agent-switcher-for-c/djflhoibgkdhkhhcedjiklpkjnoahfmg" target="_blank" rel="noopener noreferrer">User-Agent Switcher for Chrome</a>".',
    },
  };
  const i18n = (key) => i18nLib[lang][key] || `i18n[${lang}][${key}] not found`;

  if (GM_getValue('automateSearch', true)) {
    const showMsgUI = (type, msg, btnCallbackFunc) => {
      let bg;
      if (type === 'info') {
        bg = '#003aa5';
      } else if (type === 'success') {
        bg = '#00a500';
        btnCallbackFunc = () => {
          $('#mbe-bg').remove();
        };
      } else if (type === 'warn') {
        bg = '#c40000';
        btnCallbackFunc = () => {
          $('#mbe-bg').remove();
        };
      }
      $('#mbe-bg').remove();
      $('body').append(
        $(
          `<div id="mbe-bg">${msg}<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" id="mbe-msg-btn">
    <!--! Font Awesome Free 6.4.2 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. -->
    <path d="M256 48a208 208 0 1 1 0 416 208 208 0 1 1 0-416zm0 464A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM175 175c-9.4 9.4-9.4 24.6 0 33.9l47 47-47 47c-9.4 9.4-9.4 24.6 0 33.9s24.6 9.4 33.9 0l47-47 47 47c9.4 9.4 24.6 9.4 33.9 0s9.4-24.6 0-33.9l-47-47 47-47c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-47 47-47-47c-9.4-9.4-24.6-9.4-33.9 0z"></path>
  </svg>
  <style>
    #mbe-bg{position: fixed;top: 5%;left: 25%;width: 50%;padding: 10px;background: ${bg};color: #fff;border-radius: 10px;z-index: 9999;font-size: 20px;text-align: center;}
    #mbe-msg-btn{fill: #ffffff;margin-left: 10px;width: 26px;vertical-align: -5px;cursor: pointer;}}
  </style>
</div>`
        )
      );
      if (msg === i18n('gettingFromApiMsg')) {
        $('#mbe-msg-btn').remove();
      }
      $('#mbe-msg-btn').on('click', btnCallbackFunc);
    };

    const cancelSearch = () => {
      GM_setValue('automateSearchState', 0);
      GM_setValue('words', []);
      showMsgUI('warn', i18n('searchCanceledMsg'));
    };

    const inputSend = () => {
      let words = GM_getValue('words');
      let queryList;
      if (lang === 'ja') {
        queryList = [
          words[0],
          `${words[0]}の意味`,
          `${words[0]} 意味`,
          `${words[0]} とは`,
          `${words[0]} twitter`,
          `${words[0]} youtube`,
        ];
      } else {
        queryList = [
          words[0],
          `Meaning of ${words[0]}`,
          `What is the meaning of ${words[0]}?`,
          `What is ${words[0]}?`,
          `${words[0]} twitter`,
          `${words[0]} youtube`,
        ];
      }
      $('#sb_form_q').val(
        queryList[Math.floor(Math.random() * queryList.length)]
      );
      words.shift();
      if (words.length === 0) {
        GM_setValue('automateSearchState', 2);
      } else if (GM_getValue('searchesCount') === 0) {
        GM_setValue('automateSearchState', 4);
        const nowTime = new Date();
        const coolDownFinTime = [
          nowTime.getFullYear(),
          nowTime.getMonth(),
          nowTime.getDate(),
          nowTime.getHours(),
          nowTime.getMinutes(),
          nowTime.getSeconds() + unitInterval * 60,
        ];
        GM_setValue('coolDownFinTime', coolDownFinTime);
      }
      GM_setValue('words', words);
      $('#sb_form').submit();
    };

    const startAutomateSearch = (
      searchesNumPU,
      unitInterval,
      unitCount,
      iuIntervalMin,
      iuIntervalMax,
      minLength,
      maxLength
    ) => {
      GM_setValue('searchesNumPU', searchesNumPU);
      GM_setValue('unitInterval', unitInterval);
      GM_setValue('unitCount', unitCount);
      GM_setValue('iuIntervalMin', iuIntervalMin);
      GM_setValue('iuIntervalMax', iuIntervalMax);
      GM_setValue('minLength', minLength);
      GM_getValue('maxLength', maxLength);
      showMsgUI('info', i18n('gettingFromApiMsg'));
      let apiLang;
      if (lang === 'ja') {
        apiLang = 'jp';
      } else {
        apiLang = 'en';
      }
      const request = new XMLHttpRequest();
      request.open(
        'GET',
        `https://random-word.ryanrk.com/api/${apiLang}/word/random/${
          searchesNumPU * unitCount
        }?minlength=${minLength}&maxlength=${maxLength}`
      );
      request.send();
      request.addEventListener('load', function () {
        const apiData = JSON.parse(this.responseText);
        GM_setValue('words', apiData);
        GM_setValue('automateSearchState', 3);
        GM_setValue('nowUnit', 1);
        GM_setValue('searchesCount', searchesNumPU - 1);
        inputSend();
      });
    };

    let searchTimer;
    const iuIntervalMin = GM_getValue('iuIntervalMin') * 1000;
    const iuIntervalMax = GM_getValue('iuIntervalMax') * 1000;
    const searchState = GM_getValue('automateSearchState');
    const searchesNumPU = GM_getValue('searchesNumPU');
    const unitInterval = GM_getValue('unitInterval');
    const unitCount = GM_getValue('unitCount');
    const nowUnit = GM_getValue('nowUnit');
    const timeList = GM_getValue('coolDownFinTime', [2024, 0, 1, 0, 0, 0]);
    const coolDownFinTime = new Date(
      timeList[0],
      timeList[1],
      timeList[2],
      timeList[3],
      timeList[4],
      timeList[5]
    );
    if (searchState === 3) {
      showMsgUI(
        'info',
        i18n('runningMsg')
          .replace('%0', nowUnit)
          .replace('%1', unitCount)
          .replace('%2', GM_getValue('searchesCount'))
          .replace('%3', searchesNumPU)
          .replace('%4', GM_getValue('words').length)
          .replace('%5', searchesNumPU * unitCount),
        () => {
          clearTimeout(searchTimer);
          cancelSearch();
        }
      );
      GM_setValue('searchesCount', GM_getValue('searchesCount') - 1);
      searchTimer = setTimeout(
        inputSend,
        Math.floor(Math.random() * (iuIntervalMax - iuIntervalMin)) +
          iuIntervalMin
      );
    } else if (searchState === 4) {
      showMsgUI(
        'info',
        i18n('coolDownMsg')
          .replace('%0', nowUnit + 1)
          .replace('%1', unitCount)
          .replace('%2', GM_getValue('words').length)
          .replace('%3', searchesNumPU * unitCount),
        () => {
          clearInterval(searchTimer);
          cancelSearch();
        }
      );
      const countdownFunc = () => {
        const diffMSec = coolDownFinTime - new Date();
        const diffSec = Math.floor(diffMSec / 1000);
        const sec = String(diffSec % 60).padStart(2, '0');
        const min = String((diffSec - (diffSec % 60)) / 60).padStart(2, '0');
        $('#mbe-countdown').text(`${min}:${sec}`);
        if (diffMSec > 0) {
          return;
        }
        GM_setValue('automateSearchState', 3);
        GM_setValue('nowUnit', GM_getValue('nowUnit') + 1);
        GM_setValue('searchesCount', searchesNumPU - 1);
        inputSend();
      };
      searchTimer = setInterval(countdownFunc, 1000);
      countdownFunc();
    } else if (searchState === 2) {
      GM_setValue('automateSearchState', 1);
      searchTimer = setTimeout(() => {
        location.href = 'https://www.bing.com';
      }, Math.floor(Math.random() * (iuIntervalMax - iuIntervalMin)) + iuIntervalMin);
      showMsgUI(
        'info',
        i18n('runningMsg')
          .replace('%0', nowUnit)
          .replace('%1', unitCount)
          .replace('%2', GM_getValue('searchesCount'))
          .replace('%3', searchesNumPU)
          .replace('%4', GM_getValue('words').length)
          .replace('%5', searchesNumPU * unitCount),
        () => {
          clearTimeout(searchTimer);
          cancelSearch();
        }
      );
    } else if (searchState === 1) {
      GM_setValue('automateSearchState', 0);
      showMsgUI(
        'success',
        i18n('searchFinishedMsg')
          .replace('%0', unitCount)
          .replace('%1', searchesNumPU * unitCount)
      );
    }

    const showAutomateSearchUI = () => {
      let cssValLib;
      if (GM_getValue('darkMode', true)) {
        cssValLib = {
          fgBg: '#111111',
          fgColor: '#ffffff',
          inputBg: '#111111',
          inputHoverBg: '#2b2b2b',
          inputFocusBg: '#3a3a3a',
          btnBg: '#1b1b1b',
          btnHoverBg: '#2b2b2b',
        };
      } else {
        cssValLib = {
          fgBg: '#ffffff',
          fgColor: '#111111',
          inputBg: '#ffffff',
          inputHoverBg: '#dddddd',
          inputFocusBg: '#b3b3b3',
          btnBg: '#cacaca',
          btnHoverBg: '#a1a1a1',
        };
      }
      $('#mbe-bg').remove();
      let bg = $(`<div id="mbe-bg">
  <div id="mbe-fg">
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" id="mbe-close">
        <!--! Font Awesome Free 6.4.2 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. -->
        <path d="M256 48a208 208 0 1 1 0 416 208 208 0 1 1 0-416zm0 464A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM175 175c-9.4 9.4-9.4 24.6 0 33.9l47 47-47 47c-9.4 9.4-9.4 24.6 0 33.9s24.6 9.4 33.9 0l47-47 47 47c9.4 9.4 24.6 9.4 33.9 0s9.4-24.6 0-33.9l-47-47 47-47c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-47 47-47-47c-9.4-9.4-24.6-9.4-33.9 0z"></path>
    </svg>
    <h1>${i18n('automateSearch')}</h1><p class="mbe-warning">${i18n(
        'automateSearchWarning'
      )}</p>
    <p class="mbe-info">${i18n('automateSearchInfo')}</p>
    <h2>${i18n('lengthLabel')}</h2>
      <label>${i18n(
        'maxLengthLabel'
      )}<input class="mbe-input" type="number" value="${GM_getValue(
        'minLength',
        '4'
      )}" min="1" id="mbe-min-length"></label>&nbsp;<label>${i18n(
        'minLengthLabel'
      )}<input class="mbe-input" type="number" value="${GM_getValue(
        'maxLength',
        '7'
      )}" min="1" id="mbe-max-length"></label><h2>${i18n(
        'iuIntervalLabel'
      )}</h2>
      <label>${i18n(
        'iuIntervalMinLabel'
      )}<input class="mbe-input mbe-run-calc" type="number" value="${GM_getValue(
        'iuIntervalMin',
        '4'
      )}" min="1" id="mbe-iu-interval-min"></label>&nbsp;<label>${i18n(
        'iuIntervalMaxLabel'
      )}<input class="mbe-input mbe-run-calc" type="number" value="${GM_getValue(
        'iuIntervalMax',
        '7'
      )}" min="1" id="mbe-iu-interval-max"></label><br><label>${i18n(
        'searchesNumPULabel'
      )}<input class="mbe-input mbe-run-calc" type="number" value="${GM_getValue(
        'searchesNumPU',
        '3'
      )}" step="1" min="1" id="mbe-searches-num-pu"></label>&nbsp;
  <label>${i18n(
    'unitInterval'
  )}<input class="mbe-input mbe-run-calc" type="number" value="${GM_getValue(
        'unitInterval',
        '15'
      )}" step="1" min="1" id="mbe-unit-interval"></label>&nbsp;
  <label>${i18n(
    'unitCount'
  )}<input class="mbe-input mbe-run-calc" type="number" value="${GM_getValue(
        'unitCount',
        '10'
      )}" step="1" min="1" id="mbe-unit-count"></label>
    <p><span id="mbe-total-searches"></span>&nbsp;|&nbsp;${i18n(
      'estimatedTimeLabel'
    )}<span id="mbe-estimated-time"></span></p>
  <button id="mbe-search-start-btn">${i18n('searchStartBtn')}</button>
  </div>
  <style>
#mbe-bg {position: fixed;z-index: 999999;background-color: rgba(0, 0, 0, 0.8);left: 0px;top: 0px;user-select: none;-moz-user-select: none;}
#mbe-fg {width: 50%;height: 82%;padding: 15px;position: absolute;top: 10%;left: 25%;background: ${
        cssValLib['fgBg']
      };border-radius: 20px;}
#mbe-fg * {margin: 7px 0;font-family: sans-serif;font-size: 15px;color: ${
        cssValLib['fgColor']
      };}
#mbe-close {position: absolute;right: 10px;top: 10px;width: 32px;height: 32px;cursor: pointer;fill: currentColor;}
#mbe-fg p.mbe-warning{background: rgb(196, 0, 0);color: #fff;padding: 10px;border-radius: 5px;}
#mbe-fg p.mbe-info{background: rgb(0 86 196);color: #fff;padding: 10px;border-radius: 5px;}
#mbe-fg p.mbe-info a{color: #ffffff;text-decoration: underline;font-weight: bold;}
#mbe-fg label{display: inline-block;}
#mbe-fg h1{font-size: 30px;font-weight: bold;}
#mbe-fg h2{font-weight: bold;margin-bottom: -10px;}
#mbe-fg .mbe-input{background: ${
        cssValLib['inputBg']
      };border: 3px solid #666;border-radius: 5px;height: 30px;padding: 5px;font-size: 25px;width: 60px;}
#mbe-fg .mbe-input:hover{background: ${cssValLib['inputHoverBg']};}
#mbe-fg .mbe-input:focus{background: ${cssValLib['inputFocusBg']};}
#mbe-fg #mbe-search-start-btn{background: ${
        cssValLib['inputBg']
      };border: 3px solid #666;border-radius: 5px;height: 45px;width: 100%;padding: 5px;font-size: 25px;cursor: pointer;}
#mbe-fg #mbe-search-start-btn:hover{background: ${cssValLib['btnHoverBg']};
  </style>
</div>`);
      $('body').append(bg);
      $('#mbe-bg').css({
        width: document.documentElement.clientWidth + 'px',
        height: document.documentElement.clientHeight + 'px',
      });
      let resizeTimer = null;
      $(window).on('resize', () => {
        if (resizeTimer !== null) {
          clearTimeout(resizeTimer);
        }
        resizeTimer = setTimeout(() => {
          $('#mbe-bg').css({
            width: document.documentElement.clientWidth + 'px',
            height: document.documentElement.clientHeight + 'px',
          });
        }, 500);
      });
      const runCalc = () => {
        $('#mbe-total-searches').text(
          i18n('totalSearchesLabel').replace(
            '%0',
            $('#mbe-searches-num-pu').val() * $('#mbe-unit-count').val()
          )
        );
        const unitCount = $('#mbe-unit-count').val();
        const unitInterval = $('#mbe-unit-interval').val();
        const iuInterval = Math.round(
          (Number($('#mbe-iu-interval-min').val()) +
            Number($('#mbe-iu-interval-max').val())) /
            2
        );
        const searchesNumPU = $('#mbe-searches-num-pu').val();
        const estimatedSec =
          (searchesNumPU - 1) * iuInterval * unitCount +
          (unitCount - 1) * (unitInterval * 60);
        const sec = String(estimatedSec % 60).padStart(2, '0');
        const min = String(Math.floor(estimatedSec / 60) % 60).padStart(2, '0');
        const hour = String(Math.floor(estimatedSec / 3600)).padStart(2, '0');
        $('#mbe-estimated-time').text(`${hour}:${min}:${sec}`);
      };
      runCalc();
      $('.mbe-run-calc').on('input', runCalc);
      $('#mbe-close').click(function () {
        $('#mbe-bg').remove();
      });
      $('#mbe-search-start-btn').click(() => {
        startAutomateSearch(
          $('#mbe-searches-num-pu').val(),
          $('#mbe-unit-interval').val(),
          $('#mbe-unit-count').val(),
          $('#mbe-iu-interval-min').val(),
          $('#mbe-iu-interval-max').val(),
          $('#mbe-min-length').val(),
          $('#mbe-max-length').val()
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
        },
        { autoClose: false }
      );
    }

    menuId[menuId.length + 1] = GM_registerMenuCommand(
      i18n('reload'),
      () => {
        location.reload();
      },
      { autoClose: false }
    );

    return Object.freeze({
      automateSearch: menu[0][1],
      highlightAnswers: menu[1][1],
      darkMode: menu[2][1],
    });
  };
  const config = registerMenu();

  if (config.highlightAnswers) {
    let cssValLib;
    if (config.darkMode) {
      cssValLib = { correctBg: '#243b28', incorrectBg: '#853a3a' };
    } else {
      cssValLib = { correctBg: '#75f08a', incorrectBg: '#f07575' };
    }
    GM_addStyle(`/* [Microsoft Rewards] Color in the correct answer */
div#overlayWrapper[aria-label*="Microsoft Rewards"] div[iscorrectoption="True"] {
  background-color: ${cssValLib['correctBg']} !important;
}

/* [Microsoft Rewards] Color in the incorrect answers */
div#overlayWrapper[aria-label*="Microsoft Rewards"] div[iscorrectoption="False"] {
  background-color: ${cssValLib['incorrectBg']} !important;
}`);
  }
});
