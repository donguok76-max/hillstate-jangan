/**
 * 장안센트럴 관심고객 접수 → 구글시트 저장 + 대표님 카톡(알림톡)/문자/메일 알림
 * 스크립트 속성(프로젝트 설정 > 스크립트 속성):
 *   SOLAPI_KEY, SOLAPI_SECRET  : 솔라피 API 키
 *   FROM  : 솔라피에 등록한 발신번호 (숫자만)
 *   TO    : 알림 받을 대표님 휴대폰 (숫자만, 여러 명이면 쉼표)
 *   PFID  : 솔라피에 연동한 카카오 채널 pfId (알림톡 승인 후)
 *   TEMPLATE_ID : 승인된 알림톡 템플릿 ID (없으면 문자로 발송)
 *   MAIL  : 백업 알림 메일 주소
 */
var SHEET_NAME = '관심고객';

function doPost(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.website) return out_('ok');                       // 봇 함정 필드
    var name = clean_(p.name, 20), phone = String(p.phone || '').replace(/\D/g, '');
    if (!name || !/^01[016789]\d{7,8}$/.test(phone)) return out_('invalid');
    if (p.agree !== 'Y') return out_('noagree');
    var cache = CacheService.getScriptCache();
    if (cache.get('p' + phone)) return out_('dup');           // 10분 내 중복 제출 차단
    cache.put('p' + phone, '1', 600);

    var row = {
      time: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
      name: name, phone: fmt_(phone),
      type: clean_(p.type, 20), when: clean_(p.when, 20), memo: clean_(p.memo, 200),
      source: clean_(p.source, 100)
    };
    sheet_().appendRow([row.time, row.name, row.phone, row.type, row.when, row.memo, row.source, '']);
    notify_(row);
    return out_('ok');
  } catch (err) {
    console.error(err);
    return out_('error');
  }
}

function doGet() { return out_('alive'); }

function notify_(r) {
  var P = PropertiesService.getScriptProperties().getProperties();
  var text = '[장안센트럴 관심고객]\n이름: ' + r.name + '\n연락처: ' + r.phone +
    '\n희망평형: ' + (r.type || '-') + '\n통화가능: ' + (r.when || '-') +
    '\n문의: ' + (r.memo || '-') + '\n유입: ' + (r.source || '홈페이지') + '\n접수: ' + r.time;
  var sent = false;
  if (P.SOLAPI_KEY && P.SOLAPI_SECRET && P.FROM && P.TO) {
    try { sent = solapi_(P, r, text); } catch (err) { console.error('solapi', err); }
  }
  if (P.MAIL) {
    MailApp.sendEmail(P.MAIL, '[관심고객] ' + r.name + ' ' + r.phone, text + (sent ? '' : '\n\n※ 카톡/문자 발송 실패 - 확인 필요'));
  }
}

function solapi_(P, r, text) {
  var messages = P.TO.split(',').map(function (to) {
    var m = { to: to.trim(), from: P.FROM, text: text };
    if (P.PFID && P.TEMPLATE_ID) {
      m.kakaoOptions = { pfId: P.PFID, templateId: P.TEMPLATE_ID, variables: {
        '#{이름}': r.name, '#{연락처}': r.phone, '#{평형}': r.type || '-',
        '#{통화시간}': r.when || '-', '#{문의}': r.memo || '-', '#{유입}': r.source || '홈페이지', '#{접수시간}': r.time } };
    }
    return m;
  });
  var date = new Date().toISOString(), salt = Utilities.getUuid();
  var sig = Utilities.computeHmacSha256Signature(date + salt, P.SOLAPI_SECRET)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  var res = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send-many/detail', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'HMAC-SHA256 apiKey=' + P.SOLAPI_KEY + ', date=' + date + ', salt=' + salt + ', signature=' + sig },
    payload: JSON.stringify({ messages: messages })
  });
  console.log(res.getResponseCode(), res.getContentText());
  return res.getResponseCode() === 200;
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['접수시간', '이름', '연락처', '희망평형', '통화가능시간', '문의내용', '유입경로', '상담메모']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v, n) { return String(v || '').replace(/[\r\n\t]+/g, ' ').replace(/^[=+\-@]/, "'").trim().slice(0, n); }
function fmt_(d) { return d.length === 11 ? d.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3') : d.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3'); }
function out_(s) { return ContentService.createTextOutput(JSON.stringify({ result: s })).setMimeType(ContentService.MimeType.JSON); }

/** 편집기에서 직접 실행해 알림 테스트 */
function testNotify() {
  notify_({ time: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'), name: '테스트', phone: '010-0000-0000', type: '78㎡', when: '오후', memo: '테스트 접수입니다', source: '테스트' });
}
