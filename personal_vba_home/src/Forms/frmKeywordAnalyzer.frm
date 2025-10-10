VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKeywordAnalyzer 
   Caption         =   "키워드분석"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   OleObjectBlob   =   "frmKeywordAnalyzer.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmKeywordAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm1의 CommandButton1 클릭 이벤트
Private Sub CommandButton1_Click()
    ' 진텍스상품바코드정리 서브루틴 호출
    진텍스상품바코드정리
    
        ' 사용자 폼 닫기
    Unload Me
    
End Sub

Private Sub CommandButton10_Click()
    ' 볼드처리
    볼드체처리하기
     ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton11_Click()
    
    단어검색후색상으로강조하기
    ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton12_Click()

    Dim filePath As String
    filePath = "C:\Users\owner\Documents\해외주문자료\유승무역\유승무역발주서.xls"
    
    ' 파일 실행
    ' Shell "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE """ & filePath & """", vbNormalFocus
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton13_Click()
    Dim filePath As String
    filePath = "C:\Users\owner\Documents\해외주문자료\Simon_KZ_Fashions(가방바닥집)\1.Kz_orderlist_2020_0722(total).xlsx"
    
    ' 파일 실행
    ' Shell "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE """ & filePath & """", vbNormalFocus
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton14_Click()
    Dim filePath As String
    filePath = "C:\Users\owner\Documents\해외주문자료\Mr.King\bdaodao_order_2022.xlsx"
    
    ' 파일 실행
    ' Shell "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE """ & filePath & """", vbNormalFocus
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton15_Click()
    Dim filePath As String
    filePath = "C:\Users\owner\Documents\해외주문자료\PB_리사네\order_list_2020_total.xlsx"
    
    ' 파일 실행
    ' Shell "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE """ & filePath & """", vbNormalFocus
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton16_Click()
    
    Dim filePath As String
    filePath = "C:\Users\owner\Documents\python\엑셀업로드양식(스마트스토어).xlsx"
    
    ' 파일 실행
    ThisWorkbook.FollowHyperlink filePath
    
    ' 사용자 폼 닫기
    
    Unload Me
End Sub

Private Sub CommandButton17_Click()

    Dim filePath As String
    filePath = "C:\Users\owner\Documents\python\엑셀업로드양식(쿠팡).xlsx"
    
    ' 파일 실행
    ThisWorkbook.FollowHyperlink filePath
    
    ' 사용자 폼 닫기
     Unload Me
End Sub

Private Sub CommandButton18_Click()
    
    쿠팡상품정보시트_자동화
    
    ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton19_Click()

    모든이미지삭제
     ' 사용자 폼 닫기
    Unload Me
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton20_Click()

    Dim filePath As String
    filePath = "G:\내 드라이브\02_상품소싱자료\겨울상품소싱자료(모자,바라클라바,담요,양말).xlsx"
    
    ' 파일 실행
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me
        
End Sub

Private Sub CommandButton21_Click()

    Dim filePath As String
    filePath = "\\Data\e\신진우\제트배송\쿠팡상품정보\썸유로켓상품그룹핑요청목록.xlsx"
    
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton22_Click()
    쿠팡광고_키워드별_분석
    ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton23_Click()
    쿠팡광고집행상품분석
    ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton24_Click()
    '
    단어포함행삭제
     ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton25_Click()

    ' 하이퍼링크만들기
    URL생성하고하이퍼링크만들기
     ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton26_Click()
    ' 필터링_및_복사
    필터링_및_복사
     ' 사용자 폼 닫기
    Unload Me
End Sub


Private Sub CommandButton27_Click()

    ' 쿠팡광고_날짜별_캠페인별_분석

    쿠팡광고_날짜별_캠페인별_분석
     ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton28_Click()

    '쿠팡광고노출지면분석
    쿠팡광고노출지면분석
     ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton29_Click()

    '쿠팡광고전환상품분석
    쿠팡광고전환상품분석
     ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    ' 하이퍼링크만들기
    하이퍼링크만들기_열전체
     ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton30_Click()

    ' 바코드정리
    진텍스상품바코드정리
        
    ' 사용자 폼 닫기
    Unload Me

End Sub

Private Sub CommandButton31_Click()


    ' 컬럼기주색깔변경
    ColorRowsByColumnGroup
        
    ' 사용자 폼 닫기
    Unload Me
    
End Sub

Private Sub CommandButton32_Click()
    ' 다년간 키워드 분석
    RunKeywordAnalysis_MultiYear
    ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton33_Click()

    ' 중국어키워드생성
    중국어키워드생성
    ' 사용자 폼 닫기
    Unload Me
    

End Sub

Private Sub CommandButton34_Click()

    ' 급등키워드 분석
    AnalyzeKeywords_AutoPeriod
        
    ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton35_Click()

    ' 열너비 자동 조절 및 테두리치기
    데이타열너비자동조절및테두리치기
        
    ' 사용자 폼 닫기
    Unload Me
End Sub



Private Sub CommandButton36_Click()
CopyColoredCellsFromColumnB
    ' --------------------------------------------------------
    ' 프로그램명 : B열의 채우기 색 있는 셀을 B열검색어 시트로 복사
    CopyColoredCellsFromColumnB
    ' 사용자 폼 닫기
    Unload Me
    
End Sub

Private Sub CommandButton4_Click()
    ' 썸네일삽입
    썸네일삽입
    ' 사용자 폼 닫기
    Unload Me
    
End Sub

Private Sub CommandButton5_Click()
    ' 열값을입력받아숫자로전환
    열값을입력받아숫자로전환
        
    ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub CommandButton6_Click()

    Dim filePath As String
    filePath = "\\Data\e\신진우\제트배송\썸유_로켓발주서\01_로켓발주서관리프로그램_매크로사용.xlsm"
    
    ' 파일 실행
    ' Shell "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE """ & filePath & """", vbNormalFocus
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me

End Sub

Private Sub CommandButton7_Click()
    
    Dim filePath As String
    filePath = "C:\Users\owner\Documents\에스엘_재고및입고관리프로그램.xlsm"
    
    ' 파일 실행
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me

End Sub

Private Sub CommandButton8_Click()
    ' 네이버 웹사이트 열기
    ThisWorkbook.FollowHyperlink "http://www.naver.com"
    ' 사용자 폼 닫기
    Unload Me

End Sub

Private Sub CommandButton9_Click()
    Dim filePath As String
    filePath = "\\Data\e\신진우\제트배송\원가계산.xlsx"
    
    ' 파일 실행
    ThisWorkbook.FollowHyperlink filePath
    
        ' 사용자 폼 닫기
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

