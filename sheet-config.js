/*
  구글 시트를 백엔드로 쓰기 위한 설정 파일입니다.

  1. 구글 스프레드시트를 하나 새로 만드세요.
  2. 상단 메뉴 [확장 프로그램] > [Apps Script] 를 클릭하세요.
  3. 열린 편집기에 apps-script/Code.gs 파일의 내용을 그대로 붙여넣고 저장하세요.
  4. 오른쪽 위 [배포] > [새 배포] > 유형 선택(⚙️)에서 "웹 앱"을 고르세요.
     - 실행 계정: 나
     - 액세스 권한: 모든 사용자
  5. 배포 후 나오는 웹 앱 URL(.../exec 로 끝남)을 아래 SHEET_API_URL 에 붙여넣으세요.
*/

const SHEET_API_URL = "https://script.google.com/macros/s/AKfycbxonsreHCJqNktTdfv1iz72XOAsLDYW0bo9uYDkAhHzv_UjRSCA_PCLhLc3URD3-qDy/exec";
