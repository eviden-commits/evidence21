/* =========================================================================
   Auth.gs
   관리자 비밀번호 검증. 실제 값은 소스코드가 아니라 스크립트 속성
   (PropertiesService)에 저장되어 있어 git에 커밋되지 않습니다.
   ========================================================================= */

function checkPassword_(data) {
  var current = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return !!current && !!data.password && String(data.password) === current;
}

function setAdminPassword_(data) {
  var newPwd = String(data.newPassword || '').trim();
  if (!newPwd) throw new Error('새 비밀번호를 입력하세요.');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', newPwd);
}
