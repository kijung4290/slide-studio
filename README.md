# 발표 작업실 · Slide Studio

PPT·PDF 슬라이드를 왼쪽에, 발표 대본과 진행 메모를 오른쪽에 띄우는 **Windows 설치형 프로그램**입니다. 슬라이드가 바뀌면 해당 장의 대본이 함께 바뀝니다.

**[최신 설치 파일 받기](https://github.com/kijung4290/slide-studio/releases/latest)**

## 설치

1. 릴리스에서 `SlideStudio-Setup-버전-x64.exe`를 내려받아 실행합니다.
2. 설치 위치를 선택합니다. 바탕화면 또는 시작 메뉴의 **발표 작업실**로 실행합니다.
3. **파일 가져오기**에서 PPT·PPTX·PDF 또는 프로젝트 JSON을 선택합니다.

Windows 10/11 x64용입니다. Node.js나 별도 브라우저 설치 없이 실행합니다. PPT·PPTX 가져오기와 PPTX 저장에는 이 PC에 설치된 Microsoft PowerPoint가 필요합니다. PDF 가져오기, 새 슬라이드 편집과 발표는 PowerPoint 없이 사용할 수 있습니다. 초기 설치 파일은 Windows 코드 서명 인증서로 서명되지 않았습니다.

## 기능과 편집 범위

- PPT·PPTX의 글자·도형·이미지·표를 개별 요소로 가져오고, 발표자 노트를 오른쪽 대본에 연결합니다.
- 문구를 두 번 클릭해 수정하고 요소를 드래그해 이동합니다. 모서리로 크기를 조절하고, 아래 속성에서 위치·회전·글꼴·글자 크기(pt)·색을 바꿉니다.
- 표의 셀 문구 수정, 이미지 교체, 텍스트·이미지·기본 도형 추가, 요소 복제·삭제·앞뒤 배치를 지원합니다.
- **PPTX로 저장**하면 수정한 요소와 발표 대본을 PowerPoint에서 다시 편집할 수 있는 파일로 저장합니다. 원본 PPTX를 기반으로 변경한 부분을 반영합니다.
- 장별 대본·메모·계획 시간 편집, 슬라이드 추가·복제·삭제·순서 변경을 지원합니다.
- 새로 추가한 슬라이드도 글자와 도형을 직접 편집하고 PPTX로 저장할 수 있습니다.
- PDF는 페이지별 이미지로 가져오며, 텍스트·이미지를 추가하고 대본을 작성할 수 있습니다.
- 리허설, 타이머, 대본 글자 크기 조절, 대본 없는 별도 발표화면을 지원합니다.
- 프로젝트 JSON 저장·복원과 인터넷 없이 여는 발표용 HTML 내보내기를 지원합니다.

차트·SmartArt·그룹은 하나의 요소로 이동·크기 조절하며 내부 데이터 편집이나 그룹 해제는 지원하지 않습니다. 마스터에 속한 요소와 일부 특수 개체는 배경으로 표시됩니다. PowerPoint의 모든 편집 기능을 지원하지는 않습니다. 수정하지 않은 요소는 원본 서식을 유지하며, 문구나 글꼴을 수정한 텍스트 상자는 상자 전체에 기본 서식을 적용합니다. 표는 셀 문구 편집을 지원합니다.

PDF와 이전 HTML 예제는 PPTX 저장 시 배경 이미지로 들어가고 추가한 텍스트·이미지는 개별 요소로 저장됩니다. 애니메이션·동영상은 앱에서 재생하지 않습니다. 암호가 있는 파일은 암호를 해제한 뒤 가져오세요. 최대 150MB, 200장까지 지원합니다. **1.0.x에서 이미지로 가져온 PPT는 원본 PPT·PPTX를 다시 가져와야 개별 요소를 편집할 수 있습니다.**

빔 연결 시 Windows+P → 확장을 선택하세요. Ctrl+S는 프로젝트 저장, Ctrl+Shift+S는 PPTX 저장입니다. 선택한 요소는 방향키로 이동(Shift: 10씩), Enter로 문구 편집, Esc로 편집 종료, Delete로 삭제, Ctrl+D로 복제합니다. 요소가 선택되지 않은 상태에서 ← / →는 장 이동, T는 타이머, B는 발표화면 가리기, F는 전체화면입니다. 장별 계획 시간은 자동 넘김 기능이 아닙니다.

## 자동 저장과 개인정보

자료는 이 PC에서 처리합니다. 외부 변환 서비스나 AI API로 PPT·PDF·대본을 보내지 않습니다. PPT 변환 임시 파일은 변환 후 삭제합니다. PDF.js와 필요한 글꼴·CMap·WASM 리소스도 설치 파일에 포함됩니다.

자동 저장은 `%APPDATA%\SlideStudio` 아래 앱 저장 공간을 사용합니다. 업데이트·재설치에서 같은 위치를 사용하며, 제거 프로그램은 이 데이터를 지우지 않습니다. 중요한 작업은 **프로젝트 저장**으로 `.slides.json` 파일도 백업하세요. 이전 웹 버전의 자료는 웹에서 프로젝트로 저장한 뒤 설치형 앱에서 가져올 수 있습니다.

프로젝트 JSON에는 편집을 이어가기 위한 원본 PPTX와 대본·진행 메모가 포함됩니다. 발표용 HTML에는 대본·진행 메모와 원본 PPTX를 포함하지 않습니다. PPTX의 발표자 노트에는 발표 대본을 저장하며, 진행 메모는 프로젝트에만 보관합니다.

공개 저장소와 설치 파일에는 개인 발표 예제·업로드한 자료·작성한 대본·자동 저장 데이터가 포함되지 않습니다. **예제 열기**는 사용 방법을 설명하는 일반 예제를 엽니다.

## 자동 업데이트

앱 시작 후 새 버전을 확인하고 실행 중에는 4시간마다 확인합니다. 상단 **업데이트 확인** 또는 `도움말 → 업데이트 확인`으로 직접 확인할 수 있습니다.

새 버전은 자동 다운로드하며, 작업 중 재시작하지 않습니다. 앱을 종료할 때 설치하거나 **업데이트 설치**를 눌러 현재 작업을 자동 저장한 뒤 설치·재시작할 수 있습니다. 인터넷이 끊겨도 편집과 발표는 사용할 수 있습니다.

이 저장소의 공개 GitHub Releases를 사용하며 설치 파일·`.blockmap`·`latest.yml`을 함께 배포합니다. 다운로드 파일은 SHA-512로 확인합니다. 앱에 GitHub 개인 토큰을 넣지 않습니다. Windows Authenticode 서명 검증은 초기 무서명 배포에 맞춰 비활성화되어 있습니다.

## 개발과 검증

Node.js 22 이상이 필요합니다.

```powershell
git clone https://github.com/kijung4290/slide-studio.git
cd slide-studio
npm ci
npm start
```

개발 실행에서는 자동 업데이트를 하지 않습니다.

```powershell
npm run check
npm test
npm run smoke
npm run dist
```

`release` 폴더에 Windows NSIS 설치 파일을 생성합니다. `npm run smoke`는 격리한 임시 프로필에서 창·편집·저장 복원·발표화면을 확인합니다. `.qa/sample.pdf`·`.qa/sample.pptx` 파일이 있으면 가져오기와 실제 PPTX 저장·재열기도 확인합니다. `.qa/editable-rich.pptx`가 있으면 `npm test`에서 텍스트·표·그룹·복제·대본의 PPTX 왕복 검증을 수행합니다. PPT 테스트에는 PowerPoint가 필요합니다.

개발용 웹 실행은 `npm run web`입니다. 웹 버전은 `127.0.0.1:4318`, 설치형 앱 내부는 `127.0.0.1:14318`을 사용하며 외부 네트워크에서는 접근할 수 없습니다.

## 새 버전 배포

`main`에 푸시하면 GitHub Actions 검증을 실행합니다. 버전 태그를 푸시하면 Windows 설치 파일을 빌드하고 GitHub Releases에 업데이트 메타데이터와 함께 게시합니다.

```powershell
npm version patch
git push origin main --follow-tags
```

최초 배포는 현재 버전 태그를 생성합니다.

```powershell
git tag v1.0.0
git push origin v1.0.0
```

태그는 `package.json` 버전과 같아야 합니다. 워크플로는 기본 `GITHUB_TOKEN`과 `contents: write` 권한을 사용하므로 배포용 개인 토큰은 필요하지 않습니다. `latest.yml`이나 설치 파일을 삭제하면 자동 업데이트가 중단됩니다.

게시 후 실제 업데이트 확인·다운로드·해시 검증을 수행할 수 있습니다. 이 명령은 업데이트를 설치하지 않습니다.

```powershell
npm run verify:update
```

검증 결과는 Git에서 제외한 `.qa`에 저장합니다.

## 포함한 오픈소스

Electron과 electron-updater/electron-builder는 MIT, Mozilla PDF.js 6.4.299는 Apache 2.0 라이선스를 사용합니다. `vendor/pdfjs/LICENSE`에 라이선스를 포함했습니다.

공식 문서: [Electron](https://www.electronjs.org/docs/latest/tutorial/security), [electron-builder 자동 업데이트](https://www.electron.build/docs/features/auto-update/), [PDF.js](https://mozilla.github.io/pdf.js/getting_started/).
