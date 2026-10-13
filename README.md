# DAY2 — 출동 전 원인 브리핑

- 앱: `06_출동전브리핑/dist/index.html` (GitHub Pages: 저장소 루트의 `index.html`이 앱으로 이동)
- 명세: `06_출동전브리핑/SPEC.md`
- 재생성: `cd 06_출동전브리핑 && python prepare.py --mask-names`
- 테스트: `python -m unittest discover -s tests`

## 개인정보

- `site_master.xlsx`(담당자 이름·연락처)와 `조치이력.xlsx`(작업자 이름)는 `.gitignore`로 제외했다. 재생성하려면 두 파일을 원래 위치에 두어야 한다.
- 공개된 `dist/`는 이름(`박*우`)과 연락처(`010-****-1234`)를 모두 가린 빌드다.
