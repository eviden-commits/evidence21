/* =========================================================================
   Auth.gs
   관리자 비밀번호 검증.
   ========================================================================= */

function checkPassword_(data) {
  return !!data.password && String(data.password) === String(ADMIN_PASSWORD);
}
