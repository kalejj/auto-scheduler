# vendor/

이 폴더에는 외부 라이브러리를 넣습니다.

## 필요한 파일

`xlsx.full.min.js` (SheetJS) — 엑셀 다운로드 기능에 필수.

## 다운로드 (한 번만)

루트(`mobile_app/`)에서:

```bash
curl -L https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js \
  -o web/vendor/xlsx.full.min.js
```

또는 `https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js` 를
브라우저에서 열어 저장.

다운로드한 파일은 약 0.9MB. 이 폴더가 비어 있으면 엑셀 다운로드 시
"SheetJS 라이브러리가 로드되지 않았습니다" 오류가 납니다 (앱 다른 기능은 정상 작동).
