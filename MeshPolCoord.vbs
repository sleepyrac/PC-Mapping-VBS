Sub Cmd_OnLoad()
  '
  ' PC-Mapping用スクリプト
  ' ユーザメニューの定義から定義して使います。
  ' スクリプトプロシージャ　ファイル名にこれ。
  ' メッシュポリゴンの４点座標を内部属性に転記していく、スクリプト 
  ' １番目のフィールドから８番目のフィールドの順次いれていきます。
  '  x1,y1,x2,y2,x3,y3,x4,y4
  ' 走らせる前に、レイヤ名のフィールド順を確認すること。
  '
  
  ' 変数定義
  Set ObjPcmApp = CreateObject("Pcm.App")
  Set ObjPcmDocument = ObjPcmApp.GetCurrentPcmDoc
  Set ObjPcmProject = ObjPcmDocument.GetProject
  Set ObjPcmView = ObjPcmDocument.GetActiveView

  ' 下記、レイヤ名にセット
  Set ObjPcmLayer = ObjPcmProject.SearchLayer("Layer1")
  Set ObjPcmDBPol = ObjPcmLayer.GetDb(3)
  Set ObjPcmArrayLong = ObjPcmApp.CreateArrayLong
  Set ObjPcmArrayPos = ObjPcmApp.CreateArrayPos

  LongPolNum = ObjPcmLayer.GetNumOfPol(True)

  For LongCnt = 1 to LongPolNum

    BoolPol = ObjPcmLayer.PolygonIsDisable(LongCnt)

    If BoolPol = False Then

      BoolRet = ObjPcmLayer.PolygonGetPos(LongCnt, ObjPcmArrayPos, False)

      BoolRet = ObjPcmDBPol.SetCell(LongCnt, 0, ObjPcmArrayPos.GetPos(0,1,1)) 
      BoolRet = ObjPcmDBPol.SetCell(LongCnt, 1, ObjPcmArrayPos.GetPos(0,0,1)) 
      BoolRet = ObjPcmDBPol.SetCell(LongCnt, 2, ObjPcmArrayPos.GetPos(1,1,1)) 
      BoolRet = ObjPcmDBPol.SetCell(LongCnt, 3, ObjPcmArrayPos.GetPos(1,0,1)) 
      BoolRet = ObjPcmDBPol.SetCell(LongCnt, 4, ObjPcmArrayPos.GetPos(2,1,1)) 
      BoolRet = ObjPcmDBPol.SetCell(LongCnt, 5, ObjPcmArrayPos.GetPos(2,0,1)) 
      BoolRet = ObjPcmDBPol.SetCell(LongCnt, 6, ObjPcmArrayPos.GetPos(3,1,1)) 
      BoolRet = ObjPcmDBPol.SetCell(LongCnt, 7, ObjPcmArrayPos.GetPos(3,0,1)) 


    End If

  Next

End Sub
