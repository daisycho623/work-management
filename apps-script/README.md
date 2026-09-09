# Google Sheet 저장 API 연결

1. 업무대장 Google Sheet에서 **확장 프로그램 → Apps Script**를 엽니다.
2. 이 폴더의 `Code.gs` 내용을 Apps Script의 `Code.gs`에 붙여 넣고 저장합니다.
3. **배포 → 새 배포 → 웹 앱**을 선택합니다.
4. 실행 계정은 본인으로 설정하고, 액세스 사용자는 사이트 이용 범위에 맞게 선택합니다.
5. 배포 후 발급된 `/exec` URL을 프로젝트 루트 `config.js`의 `apiUrl`에 입력합니다.
6. 변경 파일을 GitHub Pages에 배포합니다.
