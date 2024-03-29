// ==UserScript==
// @name        Microsoft Bing Extender
// @name:ja     Microsoft Bing Extender
// @namespace   https://github.com/Kakejoyu/MicrosoftBingExtender
// @version     1.2.0
// @icon        https://www.bing.com/favicon.ico
// @description     User script to extend the functionality of Microsoft Bing. Supports automatic earning of Microsoft Reward search points with the latest specifications.
// @description:ja  Microsoft Bingの機能を拡張するユーザースクリプト。最新仕様のMicrosoft Reward検索ポイントの自動獲得に対応しています。
// @author      Kakejoyu
// @supportURL  https://github.com/Kakejoyu/MicrosoftBingExtender/issues
// @match       http*://www.bing.com/*
// @license     MIT
// @grant       GM.setValue
// @grant       GM.getValue
// @grant       GM.listValues
// @grant       GM.deleteValue
// @grant       GM_setValue
// @grant       GM_getValue
// @grant       GM_addStyle
// @grant       GM_listValues
// @grant       GM_deleteValue
// @require     https://cdnjs.cloudflare.com/ajax/libs/jquery/3.7.1/jquery.min.js
// @require     https://greasemonkey.github.io/gm4-polyfill/gm4-polyfill.js
// @run-at      document-end
// @noframes
// ==/UserScript==

/*! Microsoft Bing Extender | MIT license | https://github.com/Kakejoyu/MicrosoftBingExtender/blob/main/LICENSE */

jQuery(($) => {
  ('use strict');

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
      searchPausedMsg: '自動検索が一時停止されました。',
      searchRestartMsg: '自動検索を再開します。',
      setting_title: 'Microsoft Bing Extenderの設定',
      setting_switch: '機能のON/OFFの切り替え',
      setting_save_btn: '保存',
      setting_reset_btn: 'リセット',
      closeTooltip: '閉じる',
      pauseTooltip: '検索を一時停止',
      restartTooltip: '検索を再開',
      cancelTooltip: '検索を中止',
      search: '検索',
    },
    en: {
      automateSearch: 'Automate Search',
      highlightAnswers: 'Highlight Answers',
      darkMode: 'Dark Mode',
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
      searchPausedMsg: 'Automatic search has been paused.',
      searchRestartMsg: 'Restart automatic search.',
      setting_title: 'Setting of Microsoft Bing Extender',
      setting_switch: 'Switching functions ON/OFF',
      setting_save_btn: 'Save',
      setting_reset_btn: 'Reset',
      closeTooltip: 'Close',
      pauseTooltip: 'Pause search',
      restartTooltip: 'Restart search',
      cancelTooltip: 'Cancel search',
      search: 'Search',
    },
  };
  const i18n = (key) => i18nLib[lang][key] || `i18n[${lang}][${key}] not found`;

  let menuId = [];
  const initConfig = () => {
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
    }

    return Object.freeze({
      automateSearch: menu[0][1],
      highlightAnswers: menu[1][1],
      darkMode: menu[2][1],
    });
  };
  const config = initConfig();

  const callback = function (mutations, observer) {
    for (let i = 0, len = mutations.length; i < len; i++) {
      const mutation = mutations[i];
      if (mutation.type !== 'childList' || $('#mbe-setting-btn').length > 0) {
        continue;
      }

      $(mutation.target)
        .find('#HBContent > div:nth-child(3)')
        .after(
          $(
            `<div class="hb_sect_container" id="mbe-setting-btn"><div class="hb_section hb_top_sec" id="hbsettings" tabindex="0" role="menuitem" aria-haspopup="true" aria-expanded="false" aria-controls="hbsettree"><div class="hb_titlerow"><div class="hbic_col"><span class="hbic_setting"></span></div><div class="hb_title_col">${i18n(
              'setting_title'
            )}</div></div></div></div>`
          ).on('click', () => {
            let cssValLib;
            if (config.darkMode) {
              // ダークテーマ
              cssValLib = {
                fgBg: '#111111',
                fgColor: '#ffffff',
              };
            } else {
              // ライトテーマ
              cssValLib = {
                fgBg: '#ffffff',
                fgColor: '#111111',
              };
            }

            $('body').append(`<div id="mbe-bg">
    <div id="mbe-fg">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" id="mbe-close"><!--! Font Awesome Free 6.4.2 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. -->
            <path d="M256 48a208 208 0 1 1 0 416 208 208 0 1 1 0-416zm0 464A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM175 175c-9.4 9.4-9.4 24.6 0 33.9l47 47-47 47c-9.4 9.4-9.4 24.6 0 33.9s24.6 9.4 33.9 0l47-47 47 47c9.4 9.4 24.6 9.4 33.9 0s9.4-24.6 0-33.9l-47-47 47-47c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-47 47-47-47c-9.4-9.4-24.6-9.4-33.9 0z"></path>
        </svg>
        <h1>${i18n('setting_title')}</h1>
        <div id="mbe-settings">
            <h2>${i18n('setting_switch')}</h2>
            <label class="mbe-toggle-box"><input type="checkbox" name="automateSearch" /><div><div></div></div>${i18n(
              'automateSearch'
            )}</label>
            <label class="mbe-toggle-box"><input type="checkbox" name="highlightAnswers" /><div><div></div></div>${i18n(
              'highlightAnswers'
            )}</label>
            <label class="mbe-toggle-box"><input type="checkbox" name="darkMode" /><div><div></div></div>${i18n(
              'darkMode'
            )}</label>
        </div>
        <div id="mbe-btns">
            <button type="button" title="${i18n(
              'setting_save_btn'
            )}" id="mbe-save-btn">${i18n('setting_save_btn')}</button>
            <button type="button" title="${i18n(
              'setting_reset_btn'
            )}" id="mbe-reset-btn">${i18n('setting_reset_btn')}</button>
        </div>
    </div>
    <style>
        body {overflow: hidden;}
        #mbe-bg {position: fixed;z-index: 999999;background-color: rgba(0, 0, 0, 0.8);left: 0px;top: 0px;user-select: none;-moz-user-select: none;}
        #mbe-fg {width: 50%;height: 82%;padding: 15px;position: absolute;top: 10%;left: 25%;background: ${
          cssValLib['fgBg']
        };border-radius: 20px;}
        #mbe-fg * {margin: 7px 0;font-family: sans-serif;font-size: 15px;color: ${
          cssValLib['fgColor']
        };}
        #mbe-fg h1 {font-size: 25px;font-weight: bold;}
        #mbe-fg h2 {font-size: 20px;font-weight: bold;}
        #mbe-close {position: absolute;right: 10px;top: 10px;width: 32px;height: 32px;cursor: pointer;fill: currentColor;}
        #mbe-settings{height: 82%;overflow-y: scroll;}
        #mbe-fg label.mbe-toggle-box {display: block;width: fit-content;cursor: pointer;}
        #mbe-fg label.mbe-toggle-box * {margin: 0;}
        #mbe-fg label.mbe-toggle-box input {display: none;}
        #mbe-fg label.mbe-toggle-box input + div {display: inline-block;vertical-align: middle;margin-right: 10px;width: 50px;height: 24px;padding:2px;border-radius: 20px;background: #8a8a8a;position: relative;}
        #mbe-fg label.mbe-toggle-box input:checked + div {background: #0096fa;}
        #mbe-fg label.mbe-toggle-box input + div div {position: absolute;width: 24px;height: 24px;background: #ffffff;border-radius: 12px;top: 2px;left: 4%;transition: left 0.05s linear;}
        #mbe-fg label.mbe-toggle-box input:checked + div div {left: 52%;}
        #mbe-btns {display: flex;justify-content: center;}
        #mbe-btns button#mbe-save-btn, #mbe-btns button#mbe-reset-btn {font-size: 20px;width: 100px;height: 40px;border: none;border-radius: 10px;cursor: pointer;color: #ffffff;}
        #mbe-btns button#mbe-save-btn {background: #00b000;margin-right: 20px;}
        #mbe-btns button#mbe-reset-btn {background: #b00000;margin-left: 20px;}
    </style>
</div>`);
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
            $('#mbe-close').click(function () {
              $('#mbe-bg').remove();
            });

            $('#mbe-fg')
              .find('input[type="checkbox"]')
              .each(function () {
                const $checkbox = $(this);
                const name = $checkbox.attr('name');
                GM.getValue(name, true).then((value) => {
                  $checkbox.prop('checked', value);
                });
              });

            $('#mbe-fg')
              .find('#mbe-save-btn')
              .on('click', () => {
                $('#mbe-fg')
                  .find('input[type="checkbox"]')
                  .each(function () {
                    const $checkbox = $(this);
                    const name = $checkbox.attr('name');
                    const checked = $checkbox.prop('checked');
                    GM.setValue(name, checked);
                  });
                location.reload();
              });
            $('#mbe-fg')
              .find('#mbe-reset-btn')
              .on('click', () => {
                if (confirm(i18n('setting_reset_confirm'))) {
                  GM.listValues().then((keys) => {
                    keys.forEach((key) => {
                      GM.deleteValue(key);
                    });
                    location.reload();
                  });
                }
              });
          })
        );
    }
  };
  const observer = new MutationObserver(callback);
  observer.observe(document.getElementsByTagName('body')[0], {
    childList: true,
    subtree: true,
  });

  if (config.automateSearch) {
    let searchTimer;
    const showMsgUI = (bgColor, msg, btnList) => {
      $('#mbe-msg').remove();
      $('body').append(
        $(
          `<div id="mbe-msg">${msg}<span id="mbe-btn-div"></span>
  <style>
    #mbe-msg{position: fixed;top: 5%;left: 25%;width: 50%;padding: 10px;background: ${bgColor};color: #fff;border-radius: 10px;z-index: 9999;font-size: 20px;text-align: center;}
    .mbe-msg-btn{fill: #ffffff;margin-left: 10px;width: 26px;vertical-align: -5px;cursor: pointer;}}
  </style>
</div>`
        )
      );
      document.title = msg;
      btnList.forEach((btn) => {
        switch (btn) {
          case 'pause':
            $('#mbe-btn-div').append(
              $(
                `<a href="javascript:void(0);" title="${i18n(
                  'pauseTooltip'
                )}"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" class="mbe-msg-btn"><!--!Font Awesome Free 6.5.1 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2024 Fonticons, Inc.--><path d="M464 256A208 208 0 1 0 48 256a208 208 0 1 0 416 0zM0 256a256 256 0 1 1 512 0A256 256 0 1 1 0 256zm224-72V328c0 13.3-10.7 24-24 24s-24-10.7-24-24V184c0-13.3 10.7-24 24-24s24 10.7 24 24zm112 0V328c0 13.3-10.7 24-24 24s-24-10.7-24-24V184c0-13.3 10.7-24 24-24s24 10.7 24 24z"/></svg></a>`
              ).on('click', () => {
                clearTimeout(searchTimer);
                pauseSearch();
              })
            );
            break;
          case 'cancel':
            $('#mbe-btn-div').append(
              $(
                `<a href="javascript:void(0);" title="${i18n(
                  'cancelTooltip'
                )}"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" class="mbe-msg-btn"><!--!Font Awesome Free 6.5.1 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2024 Fonticons, Inc.--><path d="M464 256A208 208 0 1 0 48 256a208 208 0 1 0 416 0zM0 256a256 256 0 1 1 512 0A256 256 0 1 1 0 256zm192-96H320c17.7 0 32 14.3 32 32V320c0 17.7-14.3 32-32 32H192c-17.7 0-32-14.3-32-32V192c0-17.7 14.3-32 32-32z"/></svg></a>`
              ).on('click', () => {
                clearTimeout(searchTimer);
                cancelSearch();
              })
            );
            break;
          case 'close':
            $('#mbe-btn-div').append(
              $(
                `<a href="javascript:void(0);" title="${i18n(
                  'closeTooltip'
                )}"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" class="mbe-msg-btn"><!--!Font Awesome Free 6.5.1 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2024 Fonticons, Inc.--><path d="M256 48a208 208 0 1 1 0 416 208 208 0 1 1 0-416zm0 464A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM175 175c-9.4 9.4-9.4 24.6 0 33.9l47 47-47 47c-9.4 9.4-9.4 24.6 0 33.9s24.6 9.4 33.9 0l47-47 47 47c9.4 9.4 24.6 9.4 33.9 0s9.4-24.6 0-33.9l-47-47 47-47c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-47 47-47-47c-9.4-9.4-24.6-9.4-33.9 0z"/></svg></a>`
              ).on('click', () => {
                $('#mbe-msg').remove();
              })
            );
            break;
        }
      });
    };

    const showPauseUI = () => {
      $('#id_rh').after(
        $(
          `<a id="mbe-restart-btn" href="javascript:void(0);" title="${i18n(
            'restartTooltip'
          )}"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="width: 30px;fill: white;margin-top: 10px;cursor: pointer;"><!--!Font Awesome Free 6.5.1 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2024 Fonticons, Inc.--><path d="M464 256A208 208 0 1 0 48 256a208 208 0 1 0 416 0zM0 256a256 256 0 1 1 512 0A256 256 0 1 1 0 256zM188.3 147.1c7.6-4.2 16.8-4.1 24.3 .5l144 88c7.1 4.4 11.5 12.1 11.5 20.5s-4.4 16.1-11.5 20.5l-144 88c-7.4 4.5-16.7 4.7-24.3 .5s-12.3-12.2-12.3-20.9V168c0-8.7 4.7-16.7 12.3-20.9z"></path></svg></a>`
        ).on('click', searchRestartFunc)
      );
      $('#mbe-restart-btn').after(
        $(
          `<a id="mbe-cancel-btn" href="javascript:void(0);" title="${i18n(
            'cancelTooltip'
          )}"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="width: 30px;fill: white;margin-top: 10px;cursor: pointer;"><!--!Font Awesome Free 6.5.1 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2024 Fonticons, Inc.--><path d="M464 256A208 208 0 1 0 48 256a208 208 0 1 0 416 0zM0 256a256 256 0 1 1 512 0A256 256 0 1 1 0 256zm192-96H320c17.7 0 32 14.3 32 32V320c0 17.7-14.3 32-32 32H192c-17.7 0-32-14.3-32-32V192c0-17.7 14.3-32 32-32z"/></svg></a>`
        ).on('click', () => {
          cancelSearch();
          $('#mbe-restart-btn').remove();
          $('#mbe-cancel-btn').remove();
        })
      );
    };
    const showSearchBtn = () => {
      $('#id_rh').after(
        $(
          `<a href="javascript:void(0);" title="${i18n(
            'automateSearchOpen'
          )}"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512" style="width: 30px;fill: white;margin-top: 10px;cursor: pointer;"><!--!Font Awesome Free 6.5.1 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2024 Fonticons, Inc.--><path d="M320 0c17.7 0 32 14.3 32 32V96H472c39.8 0 72 32.2 72 72V440c0 39.8-32.2 72-72 72H168c-39.8 0-72-32.2-72-72V168c0-39.8 32.2-72 72-72H288V32c0-17.7 14.3-32 32-32zM208 384c-8.8 0-16 7.2-16 16s7.2 16 16 16h32c8.8 0 16-7.2 16-16s-7.2-16-16-16H208zm96 0c-8.8 0-16 7.2-16 16s7.2 16 16 16h32c8.8 0 16-7.2 16-16s-7.2-16-16-16H304zm96 0c-8.8 0-16 7.2-16 16s7.2 16 16 16h32c8.8 0 16-7.2 16-16s-7.2-16-16-16H400zM264 256a40 40 0 1 0 -80 0 40 40 0 1 0 80 0zm152 40a40 40 0 1 0 0-80 40 40 0 1 0 0 80zM48 224H64V416H48c-26.5 0-48-21.5-48-48V272c0-26.5 21.5-48 48-48zm544 0c26.5 0 48 21.5 48 48v96c0 26.5-21.5 48-48 48H576V224h16z"/></svg></a>`
        ).on('click', showAutomateSearchUI)
      );
    };

    const pauseSearch = () => {
      GM_setValue('beforePauseState', GM_getValue('automateSearchState'));
      GM_setValue('automateSearchState', 5);
      showPauseUI();
      showMsgUI('#FEC700', i18n('searchPausedMsg'), ['close']);
      setTimeout(() => {
        document.title = `${$('#sb_form_q').val()} - ${i18n('search')}`;
      }, 3000);
    };
    const searchRestartFunc = () => {
      $('#mbe-restart-btn').remove();
      $('#mbe-cancel-btn').remove();
      showMsgUI('#003aa5', i18n('searchRestartMsg'), []);
      if (GM_getValue('beforePauseState') === 4) {
        GM_setValue('automateSearchState', 4);
        setTimeout(cdRunFunc, 1000);
      } else {
        GM_setValue('automateSearchState', 3);
        setTimeout(() => {
          sendInput();
        }, 1000);
      }
    };

    const cancelSearch = () => {
      GM_setValue('automateSearchState', 0);
      GM_setValue('words', []);
      showMsgUI('#c40000', i18n('searchCanceledMsg'), ['close']);
      setTimeout(() => {
        document.title = `${$('#sb_form_q').val()} - ${i18n('search')}`;
      }, 3000);
      showSearchBtn();
    };

    const sendInput = () => {
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
      GM_setValue('maxLength', maxLength);
      $('#mbe-bg').remove();
      showMsgUI('#003aa5', i18n('gettingFromApiMsg'), []);
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
        sendInput();
      });
    };

    const countdownFunc = () => {
      const diffMSec = coolDownFinTime - new Date();
      const diffSec = Math.floor(diffMSec / 1000);
      const sec = String(diffSec % 60).padStart(2, '0');
      const min = String((diffSec - (diffSec % 60)) / 60).padStart(2, '0');
      $('#mbe-countdown').text(`${min}:${sec}`);
      document.title = i18n('coolDownMsg')
        .replace('<span id="mbe-countdown"></span>', `${min}:${sec}`)
        .replace('%0', nowUnit + 1)
        .replace('%1', unitCount)
        .replace('%2', GM_getValue('words').length)
        .replace('%3', searchesNumPU * unitCount);
      if (diffMSec < 60000) {
        searchTimer = setTimeout(countdownFunc, 1000);
      } else {
        searchTimer = setTimeout(countdownFunc, 30000);
      }
      if (diffMSec > 0) {
        return;
      }
      clearTimeout(searchTimer);
      GM_setValue('automateSearchState', 3);
      GM_setValue('nowUnit', GM_getValue('nowUnit') + 1);
      GM_setValue('searchesCount', searchesNumPU - 1);
      sendInput();
    };
    const cdRunFunc = () => {
      showMsgUI(
        '#009e5d',
        i18n('coolDownMsg')
          .replace('%0', nowUnit + 1)
          .replace('%1', unitCount)
          .replace('%2', GM_getValue('words').length)
          .replace('%3', searchesNumPU * unitCount),
        ['pause', 'cancel']
      );
      countdownFunc();
    };

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
    if (searchState === 5) {
      showPauseUI();
    } else if (searchState === 4) {
      cdRunFunc();
    } else if (searchState === 3) {
      showMsgUI(
        '#006d9e',
        i18n('runningMsg')
          .replace('%0', nowUnit)
          .replace('%1', unitCount)
          .replace('%2', GM_getValue('searchesCount'))
          .replace('%3', searchesNumPU)
          .replace('%4', GM_getValue('words').length)
          .replace('%5', searchesNumPU * unitCount),
        ['pause', 'cancel']
      );
      GM_setValue('searchesCount', GM_getValue('searchesCount') - 1);
      searchTimer = setTimeout(() => {
        sendInput();
      }, Math.floor(Math.random() * (iuIntervalMax - iuIntervalMin)) + iuIntervalMin);
    } else if (searchState === 2) {
      GM_setValue('automateSearchState', 1);
      searchTimer = setTimeout(() => {
        location.href = 'https://www.bing.com';
      }, Math.floor(Math.random() * (iuIntervalMax - iuIntervalMin)) + iuIntervalMin);
      showMsgUI(
        '#006d9e',
        i18n('runningMsg')
          .replace('%0', nowUnit)
          .replace('%1', unitCount)
          .replace('%2', GM_getValue('searchesCount'))
          .replace('%3', searchesNumPU)
          .replace('%4', GM_getValue('words').length)
          .replace('%5', searchesNumPU * unitCount),
        ['pause', 'cancel']
      );
    } else if (searchState === 1) {
      GM_setValue('automateSearchState', 0);
      showMsgUI(
        '#00a500',
        i18n('searchFinishedMsg')
          .replace('%0', unitCount)
          .replace('%1', searchesNumPU * unitCount),
        ['close']
      );
      setTimeout(() => {
        document.title = 'Bing';
      }, 3000);
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
    <h2>${i18n('lengthLabel')}</h2>
      <label>${i18n(
        'minLengthLabel'
      )}<input class="mbe-input" type="number" value="${GM_getValue(
        'minLength',
        '4'
      )}" min="1" id="mbe-min-length"></label>&nbsp;<label>${i18n(
        'maxLengthLabel'
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
    if (GM_getValue('automateSearchState', 0) === 0) {
      showSearchBtn();
    }
  }

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
