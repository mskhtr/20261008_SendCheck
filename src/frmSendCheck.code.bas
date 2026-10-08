'========================================================
' frmSendCheck - 送信前チェック UserForm
'
' 入れ方:
'   空の UserForm を作り、名前を frmSendCheck にする。
'   フォーム上には部品を置かない。このコードを貼る。
'   起動時に VBA が画面を自動生成する。
'
' 想定:
'   フルHD 1920 x 1080
'   Windows表示倍率 150%
'
' UserFormサイズ:
'   幅 800pt
'   高さ 460pt
'========================================================
Option Explicit

'========================================================
' 送信前チェック UserForm
'
' UserForm上には何も部品を配置しません。
' 起動時にVBAで画面を自動生成します。
'
' 想定：
'   フルHD 1920 x 1080
'   Windows表示倍率 150%
'
' UserFormサイズ：
'   幅 800pt
'   高さ 460pt
'========================================================

Private mButtonHandlers As Collection
Private mCardHandlers As Collection
Private mExpandedCard As MSForms.Frame
Private mCategoryFilter As String

Private Const FORM_W As Single = 800
Private Const FORM_H As Single = 460

Private Const MARGIN As Single = 12
Private Const GAP As Single = 10

Private Const ATTACH_H As Single = 66
Private Const ATTACH_HEADER As Single = 22
Private Const ATTACH_ROW As Single = 18
Private Const ATTACH_COLS As Long = 4

Private Sub UserForm_Initialize()

    '--------------------------------------------
    ' UserForm基本設定
    '--------------------------------------------
    Me.Caption = "送信前チェック"
    Me.Width = FORM_W
    Me.Height = FORM_H

    Me.StartUpPosition = 1       '中央表示

    Me.BackColor = RGB(244, 245, 247)

    Set mButtonHandlers = New Collection
    Set mCardHandlers = New Collection

    '--------------------------------------------
    ' 画面を作成
    '--------------------------------------------
    CreateRecipientArea

    CreateAttachmentArea

    CreateFooter

    ' マウスホイールで宛先列をスクロールする
    WheelScroll_Attach Me

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    WheelScroll_Detach

End Sub


Private Sub UserForm_Terminate()

    WheelScroll_Detach

End Sub


'========================================================
' 宛先3列
'========================================================
Private Sub CreateRecipientArea()

    Dim colWidth As Single
    Dim topPos As Single
    Dim areaHeight As Single
    Dim external As Variant
    Dim outside As Variant
    Dim internal As Variant
    Dim toCount As Long
    Dim ccCount As Long
    Dim bccCount As Long

    colWidth = 246

    external = Array( _
        Array("山田 太郎", "ABC株式会社", "営業本部第一営業部法人営業室", "課長", "taro.yamada.a1b@asahikasei.com", "To"), _
        Array("John Smith", "XYZ Corporation", "Global Sales Division Enterprise Accounts", "Manager", "john.smith.b2c@asahikasei.com", "Cc"))
    outside = Array( _
        Array("佐藤 一郎", "株式会社サンプル", "生産技術本部プロセス革新部設備技術室", "主任", "ichiro.sato.c3d@asahikasei.com", "To"), _
        Array("田中 美咲", "株式会社サンプル", "経営企画本部コーポレート統括部総務企画室", "係長", "misaki.tanaka.d4e@asahikasei.com", "Cc"), _
        Array("高橋 健", "株式会社サンプル", "財務経理本部グローバル経理部連結決算室", "一般", "ken.takahashi.e5f@asahikasei.com", "Bcc"))
    internal = Array( _
        Array("鈴木 花子", "株式会社サンプル", "情報システム本部業務基盤部システム運用室", "課長", "hanako.suzuki.f6g@asahikasei.com", "To"), _
        Array("高橋 翔", "株式会社サンプル", "研究開発本部素材科学部高分子研究室", "主任", "sho.takahashi.g7h@asahikasei.com", "Cc"), _
        Array("伊藤 健太", "株式会社サンプル", "製造本部生産管理部工程計画室", "係長", "kenta.ito.h8i@asahikasei.com", "Cc"), _
        Array("渡辺 美香", "株式会社サンプル", "品質保証本部製品安全部監査推進室", "一般", "mika.watanabe.i9j@asahikasei.com", "Bcc"), _
        Array("小林 直樹", "株式会社サンプル", "人事本部人材開発部研修企画室", "一般", "naoki.kobayashi.j0k@asahikasei.com", "Bcc"))

    CountAddressTypes external, toCount, ccCount, bccCount
    CountAddressTypes outside, toCount, ccCount, bccCount
    CountAddressTypes internal, toCount, ccCount, bccCount
    CreateRoleSummary toCount, ccCount, bccCount

    topPos = 30
    areaHeight = (Me.InsideHeight - MARGIN - 28 - GAP - ATTACH_H) - GAP - topPos

    '--------------------------------------------
    ' 社外
    '--------------------------------------------
    CreateCategoryFrame _
        "fraExternal", _
        "社外", _
        "2", _
        MARGIN, _
        topPos, _
        colWidth, _
        areaHeight, _
        RGB(176, 64, 60), _
        external

    '--------------------------------------------
    ' 部外
    '--------------------------------------------
    CreateCategoryFrame _
        "fraOutside", _
        "部外", _
        "3", _
        MARGIN + colWidth + GAP, _
        topPos, _
        colWidth, _
        areaHeight, _
        RGB(156, 112, 20), _
        outside

    '--------------------------------------------
    ' 部内
    '--------------------------------------------
    CreateCategoryFrame _
        "fraInternal", _
        "部内", _
        "5", _
        MARGIN + (colWidth + GAP) * 2, _
        topPos, _
        colWidth, _
        areaHeight, _
        RGB(42, 96, 148), _
        internal

End Sub


Private Sub CountAddressTypes( _
    ByVal recipients As Variant, _
    ByRef toCount As Long, _
    ByRef ccCount As Long, _
    ByRef bccCount As Long)

    Dim i As Long
    Dim addressType As String

    For i = LBound(recipients) To UBound(recipients)
        addressType = CStr(recipients(i)(5))

        If StrComp(addressType, "To", vbTextCompare) = 0 Then
            toCount = toCount + 1
        ElseIf StrComp(addressType, "Cc", vbTextCompare) = 0 Then
            ccCount = ccCount + 1
        ElseIf StrComp(addressType, "Bcc", vbTextCompare) = 0 Then
            bccCount = bccCount + 1
        End If
    Next i

End Sub


Private Sub CreateRoleSummary( _
    ByVal toCount As Long, _
    ByVal ccCount As Long, _
    ByVal bccCount As Long)

    Dim lbl As MSForms.Label
    Dim texts(1 To 3) As String
    Dim widths(1 To 3) As Single
    Dim i As Long
    Dim leftPos As Single

    texts(1) = "To " & toCount
    texts(2) = "Cc " & ccCount
    texts(3) = "Bcc " & bccCount
    widths(1) = 52
    widths(2) = 52
    widths(3) = 60

    leftPos = MARGIN

    For i = 1 To 3
        Set lbl = AddLabel( _
            "roleSummary_" & i, _
            texts(i), _
            leftPos, 6, widths(i), 20)

        With lbl
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .Font.Bold = True
            .ForeColor = RGB(55, 55, 60)
            .BackStyle = fmBackStyleTransparent
        End With

        leftPos = leftPos + widths(i) + 14
    Next i

End Sub


'========================================================
' カテゴリFrame
'========================================================
Private Sub CreateCategoryFrame( _
    ByVal frameName As String, _
    ByVal title As String, _
    ByVal countText As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal frameWidth As Single, _
    ByVal frameHeight As Single, _
    ByVal titleColor As Long, _
    ByVal recipients As Variant)

    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim lbl As MSForms.Label

    Dim i As Long
    Dim cardTop As Single
    Dim headerHeight As Single

    headerHeight = 34

    '--------------------------------------------
    ' 外枠。タイトルはここに置き、スクロールしない
    '--------------------------------------------
    Set fra = Me.Controls.Add( _
        "Forms.Frame.1", _
        frameName, _
        True)

    With fra
        .Left = leftPos
        .Top = topPos
        .Width = frameWidth
        .Height = frameHeight

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(226, 227, 230)
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
    End With

    '--------------------------------------------
    ' カテゴリタイトル
    '--------------------------------------------
    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        frameName & "_title", _
        True)

    With lbl
        .Left = 10
        .Top = 8
        .Width = 155
        .Height = 22

        .Caption = title

        .Font.Name = "Meiryo UI"
        .Font.Size = 11
        .Font.Bold = True

        .ForeColor = titleColor

        .BackStyle = fmBackStyleTransparent
    End With

    '人数
    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        frameName & "_count", _
        True)

    With lbl
        .Left = frameWidth - 58
        .Top = 9
        .Width = 42
        .Height = 18

        .Caption = countText

        .Font.Name = "Meiryo UI"
        .Font.Size = 8
        .Font.Bold = False

        .ForeColor = RGB(140, 140, 146)

        .TextAlign = fmTextAlignRight

        .BackStyle = fmBackStyleTransparent
    End With

    '--------------------------------------------
    ' 宛先だけをスクロールする領域
    '--------------------------------------------
    Set list = fra.Controls.Add( _
        "Forms.Frame.1", _
        frameName & "_list", _
        True)

    With list
        .Left = 1
        .Top = headerHeight
        .Width = frameWidth - 2
        .Height = frameHeight - headerHeight - 1

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsVertical
        .ScrollTop = 0
        .ScrollLeft = 0

        .KeepScrollBarsVisible = fmScrollBarsVertical
    End With

    '--------------------------------------------
    ' 宛先カード
    '--------------------------------------------
    cardTop = 6

    For i = LBound(recipients) To UBound(recipients)

        CreatePersonCard _
            list, _
            i + 1, _
            cardTop, _
            recipients(i)(0), _
            recipients(i)(1), _
            recipients(i)(2), _
            recipients(i)(3), _
            recipients(i)(4), _
            recipients(i)(5), _
            titleColor

        cardTop = cardTop + 100

    Next i

    '--------------------------------------------
    ' スクロール領域
    '--------------------------------------------
    list.ScrollHeight = cardTop + 5

    RegisterCardClick fra
    RegisterCardClick fra.Controls(frameName & "_title")
    RegisterCardClick fra.Controls(frameName & "_count")
    fra.Controls(frameName & "_title").ControlTipText = "この列だけ表示"

End Sub


'========================================================
' 人物カード
'========================================================
Private Sub CreatePersonCard( _
    ByVal parentFrame As MSForms.Frame, _
    ByVal cardNo As Long, _
    ByVal topPos As Single, _
    ByVal personName As String, _
    ByVal companyName As String, _
    ByVal departmentName As String, _
    ByVal jobTitle As String, _
    ByVal emailAddress As String, _
    ByVal addressType As String, _
    ByVal categoryColor As Long)

    Dim card As MSForms.Frame
    Dim photo As MSForms.Label
    Dim lbl As MSForms.Label

    Dim cardWidth As Single

    cardWidth = parentFrame.Width - 28

    '--------------------------------------------
    ' カード本体
    '--------------------------------------------
    Set card = parentFrame.Controls.Add( _
        "Forms.Frame.1", _
        "card_" & cardNo, _
        True)

    With card
        .Left = 8
        .Top = topPos
        .Width = cardWidth
        .Height = 92

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(226, 227, 230)
        .SpecialEffect = fmSpecialEffectFlat
    End With

    '--------------------------------------------
    ' 列の上からの番号
    '--------------------------------------------
    Set lbl = card.Controls.Add( _
        "Forms.Label.1", _
        "no_" & cardNo, _
        True)

    With lbl
        .Left = 0
        .Top = 0
        .Width = 18
        .Height = 16

        .Caption = CStr(cardNo)

        .Font.Name = "Meiryo UI"
        .Font.Size = 8
        .Font.Bold = True

        .ForeColor = RGB(255, 255, 255)
        .BackColor = categoryColor

        .TextAlign = fmTextAlignCenter

        .BorderStyle = fmBorderStyleNone
    End With

    '--------------------------------------------
    ' プロフィール写真
    '
    ' 現在は仮表示。
    ' 後でMicrosoft 365のプロフィール画像に置換する。
    '--------------------------------------------
    Set photo = card.Controls.Add( _
        "Forms.Label.1", _
        "photo_" & cardNo, _
        True)

    With photo
        .Left = 8
        .Top = 25
        .Width = 42
        .Height = 42

        .Caption = GetInitials(personName)

        .Font.Name = "Meiryo UI"
        .Font.Size = 12
        .Font.Bold = True

        .ForeColor = RGB(255, 255, 255)
        .BackColor = categoryColor

        .TextAlign = fmTextAlignCenter

        .BorderStyle = fmBorderStyleNone
    End With

    '--------------------------------------------
    ' 氏名
    '--------------------------------------------
    Set lbl = card.Controls.Add( _
        "Forms.Label.1", _
        "name_" & cardNo, _
        True)

    With lbl
        .Left = 60
        .Top = 6
        .Width = cardWidth - 158
        .Height = 16

        .Caption = personName

        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Font.Bold = True

        .ForeColor = RGB(28, 28, 30)

        .BackStyle = fmBackStyleTransparent
    End With

    '--------------------------------------------
    ' 会社名
    ' 社外だけ赤の太字
    '--------------------------------------------
    If categoryColor = RGB(176, 64, 60) Then
        AddCardDetail card, "company_" & cardNo, companyName, 60, 22, cardWidth - 68, categoryColor, True
    Else
        AddCardDetail card, "company_" & cardNo, companyName, 60, 22, cardWidth - 68, RGB(70, 70, 76)
    End If

    '--------------------------------------------
    ' 部署名
    '--------------------------------------------
    AddCardDetail card, "dept_" & cardNo, departmentName, 60, 36, cardWidth - 68, RGB(110, 110, 116)

    '--------------------------------------------
    ' 役職
    '--------------------------------------------
    AddCardDetail card, "title_" & cardNo, jobTitle, 60, 50, cardWidth - 68, RGB(110, 110, 116)

    '--------------------------------------------
    ' メールアドレス
    '--------------------------------------------
    AddCardDetail card, "mail_" & cardNo, emailAddress, 60, 64, cardWidth - 68, RGB(140, 140, 146)

    '--------------------------------------------
    ' To / CC / BCC
    ' 設定されているものだけ強調し、残りは薄くする
    '--------------------------------------------
    CreateAddressTypeMarks card, cardNo, cardWidth, addressType, categoryColor

    RegisterCardClicks card

End Sub


Private Sub AddCardDetail( _
    ByVal card As MSForms.Frame, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal width As Single, _
    ByVal foreColor As Long, _
    Optional ByVal isBold As Boolean = False)

    Dim lbl As MSForms.Label

    Set lbl = card.Controls.Add( _
        "Forms.Label.1", _
        controlName, _
        True)

    With lbl
        .Left = leftPos
        .Top = topPos
        .Width = width
        .Height = 14

        .Caption = caption

        .Font.Name = "Meiryo UI"
        .Font.Size = 8
        .Font.Bold = isBold

        .ForeColor = foreColor

        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
    End With

End Sub


Private Sub RegisterCardClicks(ByVal card As MSForms.Frame)

    Dim ctl As Object

    RegisterCardClick card

    For Each ctl In card.Controls
        If TypeName(ctl) = "Label" Then
            RegisterCardClick ctl
        End If
    Next ctl

End Sub


Private Sub RegisterCardClick(ByVal ctl As Object)

    Dim handler As CCardClick

    Set handler = New CCardClick
    Set handler.ParentForm = Me

    If TypeName(ctl) = "Frame" Then
        Set handler.CardFrame = ctl
    Else
        Set handler.CardLabel = ctl
    End If

    mCardHandlers.Add handler

End Sub


Public Sub ToggleCategoryFilter(ByVal fra As MSForms.Frame)

    If mCategoryFilter = fra.Name Then
        mCategoryFilter = ""
    Else
        mCategoryFilter = fra.Name
    End If

    ApplyCategoryFilter

End Sub


Private Sub ApplyCategoryFilter()

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim colWidth As Single
    Dim fullWidth As Single

    colWidth = 246
    fullWidth = colWidth * 3 + GAP * 2
    names = Array("fraExternal", "fraOutside", "fraInternal")

    If Not mExpandedCard Is Nothing Then
        If mCategoryFilter <> "" Then
            If mExpandedCard.Parent.Parent.Name <> mCategoryFilter Then
                ApplyCardLayout mExpandedCard, False
                Set mExpandedCard = Nothing
            End If
        End If
    End If

    For i = LBound(names) To UBound(names)

        Set fra = Me.Controls(names(i))

        If mCategoryFilter = "" Or fra.Name = mCategoryFilter Then
            fra.Visible = True

            If mCategoryFilter = "" Then
                fra.Left = MARGIN + i * (colWidth + GAP)
                SizeCategoryColumn fra, colWidth
                fra.Controls(fra.Name & "_title").ControlTipText = "この列だけ表示"
            Else
                fra.Left = MARGIN
                SizeCategoryColumn fra, fullWidth
                fra.Controls(fra.Name & "_title").ControlTipText = "3列に戻す"
            End If
        Else
            fra.Visible = False
        End If

    Next i

End Sub


Private Sub SizeCategoryColumn( _
    ByVal fra As MSForms.Frame, _
    ByVal colWidth As Single)

    Dim list As MSForms.Frame
    Dim ctl As Object
    Dim cards() As MSForms.Frame
    Dim n As Long
    Dim i As Long

    fra.Width = colWidth
    fra.Controls(fra.Name & "_count").Left = colWidth - 58

    Set list = fra.Controls(fra.Name & "_list")
    list.Width = colWidth - 2

    n = 0
    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then n = n + 1
    Next ctl

    If n = 0 Then Exit Sub

    ReDim cards(1 To n)
    i = 0

    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then
            i = i + 1
            Set cards(i) = ctl
        End If
    Next ctl

    For i = 1 To n
        cards(i).Width = list.Width - 28
        ApplyCardLayout cards(i), _
            (Not mExpandedCard Is Nothing) And (mExpandedCard Is cards(i))
    Next i

End Sub


Private Function CardIsWide(ByVal card As MSForms.Frame) As Boolean

    CardIsWide = (card.Width > 400)

End Function


Private Sub RepositionCardChrome( _
    ByVal card As MSForms.Frame, _
    ByVal cardNo As String)

    Dim leftPos As Single
    Dim kinds As Variant
    Dim i As Long
    Dim nameWidth As Single
    Dim photo As MSForms.Label
    Dim nameLabel As MSForms.Label
    Dim wide As Boolean

    wide = CardIsWide(card)

    Set photo = card.Controls("photo_" & cardNo)
    Set nameLabel = card.Controls("name_" & cardNo)

    If wide Then
        photo.Left = 16
        photo.Top = 20
        photo.Width = 72
        photo.Height = 72
        photo.Font.Size = 18

        nameLabel.Left = 104
        nameLabel.Top = 12
        nameWidth = card.Width - 206
    Else
        photo.Left = 8
        photo.Top = 25
        photo.Width = 42
        photo.Height = 42
        photo.Font.Size = 12

        nameLabel.Left = 60
        nameLabel.Top = 6
        nameWidth = card.Width - 158
    End If

    If nameWidth < 40 Then nameWidth = 40
    nameLabel.Width = nameWidth

    kinds = Array("To", "Cc", "Bcc")
    leftPos = card.Width - 94

    For i = LBound(kinds) To UBound(kinds)
        card.Controls("type_" & cardNo & "_" & kinds(i)).Left = leftPos
        leftPos = leftPos + 30
    Next i

End Sub


Public Sub ToggleCard(ByVal card As MSForms.Frame)

    If Not mExpandedCard Is Nothing Then
        If mExpandedCard Is card Then
            ApplyCardLayout card, False
            Set mExpandedCard = Nothing
            Exit Sub
        End If

        ApplyCardLayout mExpandedCard, False
    End If

    ApplyCardLayout card, True
    Set mExpandedCard = card
    ScrollCardIntoView card

End Sub


Private Sub ApplyCardLayout( _
    ByVal card As MSForms.Frame, _
    ByVal expanded As Boolean)

    Dim cardNo As String
    Dim textWidth As Single
    Dim textLeft As Single
    Dim topPos As Single
    Dim minHeight As Single
    Dim fields(1 To 4) As String
    Dim tops(1 To 4) As Single
    Dim i As Long
    Dim lbl As MSForms.Label
    Dim wide As Boolean

    cardNo = Mid(card.Name, 6)
    wide = CardIsWide(card)
    RepositionCardChrome card, cardNo

    If wide Then
        textLeft = 104
        textWidth = card.Width - textLeft - 12
        minHeight = 104
        tops(1) = 32
        tops(2) = 50
        tops(3) = 68
        tops(4) = 86
    Else
        textLeft = 60
        textWidth = card.Width - 68
        minHeight = 92
        tops(1) = 22
        tops(2) = 36
        tops(3) = 50
        tops(4) = 64
    End If

    fields(1) = "company_" & cardNo
    fields(2) = "dept_" & cardNo
    fields(3) = "title_" & cardNo
    fields(4) = "mail_" & cardNo

    If expanded Then
        topPos = tops(1)

        For i = 1 To 4
            Set lbl = card.Controls(fields(i))
            lbl.Left = textLeft
            lbl.Top = topPos
            lbl.WordWrap = True
            lbl.AutoSize = False
            lbl.Width = textWidth
            lbl.Height = WrappedLineCount(lbl, textWidth) * 14
            topPos = lbl.Top + lbl.Height + 2
        Next i

        card.Height = topPos + 6
        If card.Height < minHeight Then card.Height = minHeight
    Else
        For i = 1 To 4
            Set lbl = card.Controls(fields(i))
            lbl.AutoSize = False
            lbl.WordWrap = False
            lbl.Left = textLeft
            lbl.Top = tops(i)
            lbl.Width = textWidth
            lbl.Height = 14
        Next i

        card.Height = minHeight
    End If

    ReflowCards card.Parent

End Sub


Private Function WrappedLineCount( _
    ByVal lbl As MSForms.Label, _
    ByVal maxWidth As Single) As Long

    Dim fullWidth As Single

    lbl.WordWrap = False
    lbl.AutoSize = True
    fullWidth = lbl.Width
    lbl.AutoSize = False
    lbl.WordWrap = True
    lbl.Width = maxWidth

    If maxWidth <= 0 Then
        WrappedLineCount = 1
    Else
        WrappedLineCount = -Int(-fullWidth / maxWidth)
    End If

    If WrappedLineCount < 1 Then WrappedLineCount = 1

End Function


Private Sub ReflowCards(ByVal list As MSForms.Frame)

    Dim ctl As Object
    Dim cards() As MSForms.Frame
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As MSForms.Frame
    Dim topPos As Single

    n = 0
    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then
            n = n + 1
        End If
    Next ctl

    If n = 0 Then Exit Sub

    ReDim cards(1 To n)
    i = 0
    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then
            i = i + 1
            Set cards(i) = ctl
        End If
    Next ctl

    For i = 1 To n - 1
        For j = i + 1 To n
            If CLng(Mid(cards(j).Name, 6)) < CLng(Mid(cards(i).Name, 6)) Then
                Set tmp = cards(i)
                Set cards(i) = cards(j)
                Set cards(j) = tmp
            End If
        Next j
    Next i

    topPos = 6
    For i = 1 To n
        cards(i).Top = topPos
        topPos = topPos + cards(i).Height + 8
    Next i

    list.ScrollHeight = topPos + 5

End Sub


Private Sub ScrollCardIntoView(ByVal card As MSForms.Frame)

    Dim list As MSForms.Frame
    Dim viewHeight As Single
    Dim newTop As Single

    Set list = card.Parent
    viewHeight = list.InsideHeight
    If viewHeight <= 0 Then viewHeight = list.Height

    If card.Height >= viewHeight Then
        list.ScrollTop = card.Top
        Exit Sub
    End If

    If card.Top < list.ScrollTop Then
        list.ScrollTop = card.Top
    ElseIf card.Top + card.Height > list.ScrollTop + viewHeight Then
        newTop = card.Top + card.Height - viewHeight
        If newTop < 0 Then newTop = 0
        list.ScrollTop = newTop
    End If

End Sub


Private Sub CreateAddressTypeMarks( _
    ByVal card As MSForms.Frame, _
    ByVal cardNo As Long, _
    ByVal cardWidth As Single, _
    ByVal addressType As String, _
    ByVal categoryColor As Long)

    Dim kinds(1 To 3) As String
    Dim i As Long
    Dim lbl As MSForms.Label
    Dim markWidth As Single
    Dim gap As Single
    Dim leftPos As Single
    Dim selected As Boolean

    kinds(1) = "To"
    kinds(2) = "Cc"
    kinds(3) = "Bcc"

    markWidth = 28
    gap = 2
    leftPos = cardWidth - (markWidth * 3 + gap * 2) - 6

    For i = 1 To 3

        selected = (StrComp(addressType, kinds(i), vbTextCompare) = 0)

        Set lbl = card.Controls.Add( _
            "Forms.Label.1", _
            "type_" & cardNo & "_" & kinds(i), _
            True)

        With lbl
            .Left = leftPos
            .Top = 5
            .Width = markWidth
            .Height = 16

            .Caption = kinds(i)

            .Font.Name = "Meiryo UI"
            .Font.Bold = selected

            If selected Then
                .Font.Size = 9
            Else
                .Font.Size = 7
            End If

            .TextAlign = fmTextAlignCenter

            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone

            If selected Then
                .ForeColor = categoryColor
            Else
                .ForeColor = RGB(196, 196, 200)
            End If
        End With

        leftPos = leftPos + markWidth + gap

    Next i

End Sub


'========================================================
' 添付ファイル
'========================================================
Private Sub CreateAttachmentArea()

    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim lbl As MSForms.Label
    Dim files As Variant
    Dim fileCount As Long
    Dim i As Long
    Dim col As Long
    Dim row As Long
    Dim fileWidth As Single
    Dim gapX As Single
    Dim pad As Single
    Dim contentHeight As Single

    files = Array( _
        "見積書.xlsx", _
        "提案資料.pdf", _
        "契約書.docx", _
        "議事録.docx", _
        "仕様書.xlsx", _
        "案内状.pdf", _
        "写真.jpg", _
        "名簿.xlsx", _
        "日程表.xlsx", _
        "補足.txt")

    fileCount = UBound(files) - LBound(files) + 1
    gapX = 12
    pad = 8

    Set fra = Me.Controls.Add( _
        "Forms.Frame.1", _
        "fraAttachments", _
        True)

    With fra
        .Left = MARGIN
        .Top = Me.InsideHeight - MARGIN - 28 - GAP - ATTACH_H
        .Width = 246 * 3 + GAP * 2
        .Height = ATTACH_H

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(226, 227, 230)
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
    End With

    'タイトル
    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        "lblAttachmentTitle", _
        True)

    With lbl
        .Left = 10
        .Top = 2
        .Width = 120
        .Height = 18

        .Caption = "添付ファイル"

        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Font.Bold = True

        .ForeColor = RGB(140, 140, 146)

        .BackStyle = fmBackStyleTransparent
    End With

    '件数
    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        "lblAttachmentCount", _
        True)

    With lbl
        .Left = fra.Width - 58
        .Top = 4
        .Width = 42
        .Height = 16

        .Caption = fileCount & "件"

        .Font.Name = "Meiryo UI"
        .Font.Size = 8
        .Font.Bold = False

        .ForeColor = RGB(140, 140, 146)

        .TextAlign = fmTextAlignRight

        .BackStyle = fmBackStyleTransparent
    End With

    '--------------------------------------------
    ' ファイル名だけをスクロールする領域
    '--------------------------------------------
    Set list = fra.Controls.Add( _
        "Forms.Frame.1", _
        "fraAttachments_list", _
        True)

    With list
        .Left = 1
        .Top = ATTACH_HEADER
        .Width = fra.Width - 2
        .Height = ATTACH_H - ATTACH_HEADER - 1

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
        .ScrollTop = 0
        .ScrollLeft = 0
    End With

    fileWidth = (list.Width - pad * 2 - gapX * (ATTACH_COLS - 1) - 16) / ATTACH_COLS

    For i = LBound(files) To UBound(files)

        col = (i - LBound(files)) Mod ATTACH_COLS
        row = (i - LBound(files)) \ ATTACH_COLS

        CreateAttachmentLabel _
            list, _
            i - LBound(files) + 1, _
            CStr(files(i)), _
            pad + col * (fileWidth + gapX), _
            2 + row * ATTACH_ROW, _
            fileWidth

    Next i

    contentHeight = ((fileCount + ATTACH_COLS - 1) \ ATTACH_COLS) * ATTACH_ROW + 6
    list.ScrollHeight = contentHeight

    If contentHeight > list.Height Then
        list.ScrollBars = fmScrollBarsVertical
        list.KeepScrollBarsVisible = fmScrollBarsVertical
    End If

End Sub


Private Sub CreateAttachmentLabel( _
    ByVal parentFrame As MSForms.Frame, _
    ByVal fileNo As Long, _
    ByVal fileName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal labelWidth As Single)

    Dim lbl As MSForms.Label

    Set lbl = parentFrame.Controls.Add( _
        "Forms.Label.1", _
        "file_" & fileNo, _
        True)

    With lbl
        .Left = leftPos
        .Top = topPos
        .Width = labelWidth
        .Height = 16

        .Caption = fileName

        .Font.Name = "Meiryo UI"
        .Font.Size = 8
        .Font.Bold = False

        .ForeColor = RGB(55, 55, 60)

        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone

        .WordWrap = False
    End With

End Sub


'========================================================
' フッター
'========================================================
Private Sub CreateFooter()

    Dim btnCancel As MSForms.CommandButton
    Dim btnSend As MSForms.CommandButton

    'キャンセル
    Set btnCancel = Me.Controls.Add( _
        "Forms.CommandButton.1", _
        "btnCancel", _
        True)

    With btnCancel
        .Left = Me.InsideWidth - MARGIN - 82 - 10 - 82
        .Top = Me.InsideHeight - MARGIN - 28
        .Width = 82
        .Height = 28

        .Caption = "キャンセル"

        .Font.Name = "Meiryo UI"
        .Font.Size = 9
    End With

    RegisterButton btnCancel, "Cancel"

    '送信
    Set btnSend = Me.Controls.Add( _
        "Forms.CommandButton.1", _
        "btnSend", _
        True)

    With btnSend
        .Left = Me.InsideWidth - MARGIN - 82
        .Top = Me.InsideHeight - MARGIN - 28
        .Width = 82
        .Height = 28

        .Caption = "送信する"

        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Font.Bold = True
    End With

    RegisterButton btnSend, "Send"

End Sub


'========================================================
' Buttonイベント登録
'========================================================
Private Sub RegisterButton( _
    ByVal button As MSForms.CommandButton, _
    ByVal actionName As String)

    Dim handler As CButtonHandler

    Set handler = New CButtonHandler

    Set handler.Button = button
    Set handler.ParentForm = Me
    handler.ActionName = actionName

    mButtonHandlers.Add handler

End Sub


'========================================================
' ボタンから呼ばれる処理
'========================================================
Public Sub ButtonAction(ByVal actionName As String)

    Select Case actionName

        Case "Cancel"

            Unload Me

        Case "Send"

            MsgBox _
                "これは試作版です。" & vbCrLf & vbCrLf & _
                "次の段階で、実際のOutlookメールを送信する処理を追加します。", _
                vbInformation, _
                "送信前チェック"

    End Select

End Sub


'========================================================
' イニシャル取得
'========================================================
Private Function GetInitials(ByVal fullName As String) As String

    Dim parts() As String

    fullName = Trim(fullName)

    If InStr(fullName, " ") > 0 Then

        parts = Split(fullName, " ")

        If UBound(parts) >= 1 Then
            GetInitials = Left(parts(0), 1) & Left(parts(1), 1)
        Else
            GetInitials = Left(fullName, 2)
        End If

    Else

        GetInitials = Left(fullName, 2)

    End If

End Function


'========================================================
' Label作成ヘルパー
'========================================================
Private Function AddLabel( _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal width As Single, _
    ByVal height As Single) As MSForms.Label

    Dim lbl As MSForms.Label

    Set lbl = Me.Controls.Add( _
        "Forms.Label.1", _
        controlName, _
        True)

    With lbl
        .Left = leftPos
        .Top = topPos
        .Width = width
        .Height = height
        .Caption = caption
    End With

    Set AddLabel = lbl

End Function
