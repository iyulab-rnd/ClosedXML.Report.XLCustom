[라이브러리 작성지침]
- ClosedXML.Report는 모든변수가 확정적일 것을 가정하고 있기 때문에 ""ClosedXML.Report의 모든 문법을 계승"하면서 사용자가 format, function, variable resolver 세가지 기능으로 프로그래밍 방식으로 커스터마이징 기능을 제공하는것이 목표
- ClosedXML, ClosedXML.Report 에 내장된 기능을 최대한 활용, 위임 하고 최대한 얇은 래퍼가 되도록 함.
- 범용 라이브러리로 하드코딩으로 작성하지 않고, 구성가능하도록 작성
- 코드, 주석은 영문으로 작성
- 주석은 <summary> 태그만 작성하고 <params>,<return> 등 다른 태그는 작성하지 않음
- Helpers, Extensions 등 공통함수를 사용하고 간결하게 코드 작성
- 코드가 500줄을 넘어가면 partial classs 로 {ClassName}.{Regions}.cs 으로 나누어 작성
- 개발 단계로 하위호환을 고려하지 않음, 문제해결을 위해 필요한 경우 억지로 해결하지말고 더 나은 방법을 제안

(답변은 한국어로)