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
Private mAttachHandlers As Collection
Private mScrollDrags As Collection
Private mExpandedCard As MSForms.Frame
Private mCategoryFilter As String

Private Type AttachItem
    Id As String
    ParentId As String
    Caption As String
    Status As String
    Kind As String
    Depth As Long
    Expandable As Boolean
    PasswordState As String
End Type

Private mAttachItems() As AttachItem
Private mAttachCount As Long
Private mCollapsedIds As Collection
Private mAttachExpanded As Boolean

Private Const FORM_W As Single = 800
Private Const FORM_H As Single = 460

Private Const MARGIN As Single = 12
Private Const GAP As Single = 10

Private Const ATTACH_H As Single = 78
Private Const ATTACH_HEADER As Single = 22
Private Const ATTACH_ROW As Single = 22

' リストの右に、スクロールバー専用の溝を空ける
Private Const SCROLL_LANE As Single = 10

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
    Set mAttachHandlers = New Collection
    Set mScrollDrags = New Collection
    Set mCollapsedIds = New Collection

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
        .Width = frameWidth - 2 - SCROLL_LANE
        .Height = frameHeight - headerHeight - 1

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
        .ScrollTop = 0
        .ScrollLeft = 0
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
    SyncScrollThumb list

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

    cardWidth = parentFrame.Width - 16

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
    list.Width = colWidth - 2 - SCROLL_LANE

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
        cards(i).Width = list.Width - 16
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
            lbl.Visible = True
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
    SyncScrollThumb list

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
        SyncScrollThumb list
        Exit Sub
    End If

    If card.Top < list.ScrollTop Then
        list.ScrollTop = card.Top
    ElseIf card.Top + card.Height > list.ScrollTop + viewHeight Then
        newTop = card.Top + card.Height - viewHeight
        If newTop < 0 Then newTop = 0
        list.ScrollTop = newTop
    End If

    SyncScrollThumb list

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
' ダミー。ZIP の行を押すと、その中身だけ開く。
'========================================================
Private Sub CreateAttachmentArea()

    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim lbl As MSForms.Label

    LoadDummyAttachments

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

    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        "lblAttachmentTitle", _
        True)

    With lbl
        .Left = 10
        .Top = 2
        .Width = 220
        .Height = 18

        .Caption = ChrW(&H25B8) & "  添付ファイル"
        .ControlTipText = "クリックで一覧を上まで広げる"

        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Font.Bold = True

        .ForeColor = RGB(140, 140, 146)

        .BackStyle = fmBackStyleTransparent
    End With

    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        "lblAttachmentAlert", _
        True)

    With lbl
        .Left = 130
        .Top = 3
        .Width = 280
        .Height = 16

        .Caption = "要確認の添付ファイルあり"
        .Visible = False

        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Font.Bold = True

        .ForeColor = RGB(176, 64, 60)

        .BackStyle = fmBackStyleTransparent
    End With

    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        "lblAttachmentCount", _
        True)

    With lbl
        .Left = fra.Width - 58
        .Top = 4
        .Width = 42
        .Height = 16

        .Caption = LeafFileCount() & "件"

        .Font.Name = "Meiryo UI"
        .Font.Size = 8
        .Font.Bold = False

        .ForeColor = RGB(140, 140, 146)

        .TextAlign = fmTextAlignRight

        .BackStyle = fmBackStyleTransparent
    End With

    Set list = fra.Controls.Add( _
        "Forms.Frame.1", _
        "fraAttachments_list", _
        True)

    With list
        .Left = 1
        .Top = ATTACH_HEADER
        .Width = fra.Width - 2 - SCROLL_LANE
        .Height = ATTACH_H - ATTACH_HEADER - 1

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
        .ScrollTop = 0
        .ScrollLeft = 0
    End With

    RegisterCardClick fra
    RegisterCardClick fra.Controls("lblAttachmentTitle")
    RegisterCardClick fra.Controls("lblAttachmentAlert")
    RegisterCardClick fra.Controls("lblAttachmentCount")
    RegisterCardClick list

    RenderAttachmentList

End Sub


Public Sub ToggleAttachment(ByVal attachId As String)

    Dim index As Long

    index = FindAttachItem(attachId)
    If index = 0 Then Exit Sub
    If Not mAttachItems(index).Expandable Then Exit Sub

    If AttachItemIsExpanded(attachId) Then
        CollapseAttachment attachId
    Else
        ExpandAttachment attachId
    End If

    RenderAttachmentList

End Sub


Private Sub LoadDummyAttachments()

    mAttachCount = 0
    ExpandAllAttachments

    ' PDF だけ
    AddAttachItem "a1", "", "見積.pdf", "PDF", "pdf", 0, False

    ' 拡張子が変わった ZIP。中は PDF だけ。パスワードは未設定
    AddAttachItem "a2", "", "資料.zi_", "中は PDF  3件", "zipok", 0, True, "none"
    AddAttachItem "a2a", "a2", "表紙.pdf", "PDF", "pdf", 1, False
    AddAttachItem "a2b", "a2", "本編.pdf", "PDF", "pdf", 1, False
    AddAttachItem "a2c", "a2", "別紙.pdf", "PDF", "pdf", 1, False

    ' 入れ子 ZIP。器は判定に出さない。中に PDF 以外がある。パスワードは未設定
    AddAttachItem "a3", "", "提案.zip", "PDF以外あり", "zipbad", 0, True, "none"
    AddAttachItem "a3a", "a3", "cover.pdf", "PDF", "pdf", 1, False
    AddAttachItem "a3b", "a3", "内訳.zip", "器", "container", 1, True, "none"
    AddAttachItem "a3b1", "a3b", "見積.pdf", "PDF", "pdf", 2, False
    AddAttachItem "a3b2", "a3b", "仕様.docx", "PDF以外", "other", 2, False

    ' パスワード付きだが、中身は確認できない
    AddAttachItem "a4", "", "契約.zip", "", "password", 0, True, "set"
    AddAttachItem "a4n", "a4", "中身は確認できません", "", "note", 1, False

    ' パスワード付き ZIP の中の PDF。これだけ安全
    AddAttachItem "a5", "", "送付.zip", "", "password", 0, True, "set"
    AddAttachItem "a5a", "a5", "本文.pdf", "PDF", "pdf", 1, False

End Sub


Private Sub AddAttachItem( _
    ByVal itemId As String, _
    ByVal parentId As String, _
    ByVal caption As String, _
    ByVal statusText As String, _
    ByVal kind As String, _
    ByVal depth As Long, _
    ByVal expandable As Boolean, _
    Optional ByVal passwordState As String = "")

    mAttachCount = mAttachCount + 1
    ReDim Preserve mAttachItems(1 To mAttachCount)

    With mAttachItems(mAttachCount)
        .Id = itemId
        .ParentId = parentId
        .Caption = caption
        .Status = statusText
        .Kind = kind
        .Depth = depth
        .Expandable = expandable
        .PasswordState = passwordState
    End With

End Sub


Private Sub RenderAttachmentList()

    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim visibleIds() As String
    Dim n As Long
    Dim i As Long
    Dim contentH As Single
    Dim collapsedContentH As Single
    Dim topCount As Long
    Dim attachH As Single
    Dim maxH As Single
    Dim listH As Single
    Dim savedScroll As Single

    Set fra = Me.Controls("fraAttachments")
    Set list = fra.Controls("fraAttachments_list")
    savedScroll = list.ScrollTop

    Set mAttachHandlers = New Collection
    ClearFrameControls list

    n = 0
    ReDim visibleIds(1 To mAttachCount)

    For i = 1 To mAttachCount
        If AttachItemVisible(i) Then
            n = n + 1
            visibleIds(n) = mAttachItems(i).Id
        End If
    Next i

    contentH = n * ATTACH_ROW + 4
    topCount = 0
    For i = 1 To mAttachCount
        If Len(mAttachItems(i).ParentId) = 0 Then topCount = topCount + 1
    Next i
    collapsedContentH = topCount * ATTACH_ROW + 4

    If mAttachExpanded Then
        fra.Top = 30
        attachH = Me.InsideHeight - MARGIN - 28 - GAP - fra.Top
    Else
        maxH = Me.InsideHeight - MARGIN - 28 - GAP - 30 - GAP - 150

        If ATTACH_HEADER + collapsedContentH + 2 > maxH Then
            attachH = maxH
        Else
            attachH = ATTACH_HEADER + collapsedContentH + 2
        End If

        ' 添付欄だけが高くなりすぎないよう、中身に合わせた高さの 2/3 にする
        attachH = attachH * 2 / 3
    End If

    If attachH < ATTACH_HEADER + ATTACH_ROW + 6 Then
        attachH = ATTACH_HEADER + ATTACH_ROW + 6
    End If

    If Not mAttachExpanded Then
        fra.Top = Me.InsideHeight - MARGIN - 28 - GAP - attachH
    End If

    listH = attachH - ATTACH_HEADER - 1

    fra.Height = attachH
    list.Height = listH

    list.ScrollBars = fmScrollBarsNone
    list.ScrollHeight = contentH

    For i = 1 To n
        If i > 1 Then
            If Len(mAttachItems(FindAttachItem(visibleIds(i))).ParentId) = 0 Then
                AddAttachSeparator list, visibleIds(i), (1 + (i - 1) * ATTACH_ROW) - 1, list.InsideWidth - 4
            End If
        End If
        CreateAttachRow list, visibleIds(i), 1 + (i - 1) * ATTACH_ROW, list.InsideWidth - 4
    Next i

    If n > 0 Then
        AddAttachSeparator list, "top", 0, list.InsideWidth - 4
        AddAttachSeparator list, "bottom", n * ATTACH_ROW, list.InsideWidth - 4
    End If

    If savedScroll < 0 Then savedScroll = 0
    If list.ScrollHeight > list.InsideHeight Then
        If savedScroll > list.ScrollHeight - list.InsideHeight Then
            savedScroll = list.ScrollHeight - list.InsideHeight
        End If
        list.ScrollTop = savedScroll
    Else
        list.ScrollTop = 0
    End If

    If mAttachExpanded Then
        HideRecipientColumns
    Else
        ResizeRecipientColumns fra.Top
    End If

    UpdateAttachmentHeading
    SyncScrollThumb list

End Sub


Public Sub ToggleAttachmentFocus()

    mAttachExpanded = Not mAttachExpanded

    If mAttachExpanded Then
        ExpandAllAttachments
        HideRecipientColumns
    Else
        ApplyCategoryFilter
    End If

    RenderAttachmentList

End Sub


Private Sub HideRecipientColumns()

    Dim names As Variant
    Dim i As Long

    names = Array("fraExternal", "fraOutside", "fraInternal")

    For i = LBound(names) To UBound(names)
        Me.Controls(CStr(names(i))).Visible = False
    Next i

End Sub


Private Sub UpdateAttachmentHeading()

    Dim title As MSForms.Label
    Dim alert As MSForms.Label
    Dim titleWidth As Single

    Set title = Me.Controls("fraAttachments").Controls("lblAttachmentTitle")
    Set alert = Me.Controls("fraAttachments").Controls("lblAttachmentAlert")

    If mAttachExpanded Then
        title.Caption = ChrW(&H25BE) & "  添付ファイル"
        title.ControlTipText = "クリックで宛先カードを戻す"
    Else
        title.Caption = ChrW(&H25B8) & "  添付ファイル"
        title.ControlTipText = "クリックで一覧を上まで広げる"
    End If

    titleWidth = MeasureLabelWidth(title)
    title.AutoSize = False
    title.Width = titleWidth + 2
    title.Height = 18
    title.Top = 2

    alert.Visible = HasAttachmentToReview()
    If Not alert.Visible Then Exit Sub

    alert.AutoSize = False
    alert.Width = MeasureLabelWidth(alert) + 4
    alert.AutoSize = False
    alert.Height = 16
    alert.Left = title.Left + title.Width + 10
    alert.Top = 3

End Sub


Private Function HasAttachmentToReview() As Boolean

    Dim i As Long

    For i = 1 To mAttachCount
        If Not AttachmentIsSafe(i) Then
            HasAttachmentToReview = True
            Exit Function
        End If
    Next i

End Function


Private Function LeafFileCount() As Long

    Dim i As Long

    For i = 1 To mAttachCount
        If mAttachItems(i).Expandable Then
            ' ZIP と空の ZIP は件数に入れない
        ElseIf mAttachItems(i).Kind <> "note" Then
            LeafFileCount = LeafFileCount + 1
        End If
    Next i

End Function


Private Function AttachmentIsSafe(ByVal index As Long) As Boolean

    Dim parentIndex As Long
    Dim sawPasswordZip As Boolean

    If mAttachItems(index).Kind <> "pdf" Then Exit Function

    parentIndex = FindAttachItem(mAttachItems(index).ParentId)

    Do While parentIndex > 0
        If mAttachItems(parentIndex).PasswordState = "none" Then Exit Function
        If mAttachItems(parentIndex).PasswordState = "set" Then sawPasswordZip = True
        parentIndex = FindAttachItem(mAttachItems(parentIndex).ParentId)
    Loop

    AttachmentIsSafe = sawPasswordZip

End Function


Private Sub AddAttachSeparator( _
    ByVal list As MSForms.Frame, _
    ByVal itemId As String, _
    ByVal topPos As Single, _
    ByVal rowWidth As Single)

    Dim line As MSForms.Label

    Set line = list.Controls.Add( _
        "Forms.Label.1", _
        "ln_" & itemId, _
        True)

    With line
        .Left = 0
        .Top = topPos
        .Width = rowWidth
        .Height = 1
        .Caption = ""
        .BackColor = RGB(186, 188, 194)
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
    End With

    line.ZOrder 0

End Sub


Private Sub CreateAttachRow( _
    ByVal list As MSForms.Frame, _
    ByVal itemId As String, _
    ByVal topPos As Single, _
    ByVal rowWidth As Single)

    Dim index As Long
    Dim lbl As MSForms.Label
    Dim nameColor As Long
    Dim statusColor As Long
    Dim backColor As Long
    Dim nameLeft As Single
    Dim nameWidth As Single
    Dim statusWidth As Single
    Dim mark As String
    Dim markLeft As Single
    Dim markLbl As MSForms.Label
    Dim nameLbl As MSForms.Label
    Dim statusLbl As MSForms.Label
    Dim pwLbl As MSForms.Label
    Dim cursorLeft As Single
    Dim tailRight As Single

    index = FindAttachItem(itemId)
    If index = 0 Then Exit Sub

    AttachRowStyle index, nameColor, statusColor, backColor

    ' 実行時に作った Frame の中へは部品を置けないので、行はリストへ直接並べる
    Set lbl = list.Controls.Add( _
        "Forms.Label.1", _
        "row_" & itemId, _
        True)

    With lbl
        .Left = 0
        .Top = topPos
        .Width = rowWidth
        .Height = ATTACH_ROW - 1
        .Caption = ""
        .BackColor = backColor
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .WordWrap = False
    End With

    lbl.ZOrder 1
    RegisterAttachClick lbl

    ' 左端に余白を取り、ファイル名は赤線を引いたあたりから始める
    markLeft = 26 + mAttachItems(index).Depth * 16
    nameLeft = markLeft + 14

    If mAttachItems(index).Expandable Then
        If AttachItemIsExpanded(itemId) Then
            mark = ChrW(&H25BE)
        Else
            mark = ChrW(&H25B8)
        End If

        Set markLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "mk_" & itemId, _
            True)

        With markLbl
            .Left = markLeft
            .Top = topPos + 2
            .Width = 14
            .Height = 16
            .Caption = mark
            .Font.Name = "Meiryo UI"
            .Font.Size = 9
            .ForeColor = nameColor
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        markLbl.ZOrder 0
        RegisterAttachClick markLbl
    ElseIf mAttachItems(index).Kind <> "note" Then
        Set markLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "fi_" & itemId, _
            True)

        With markLbl
            .Left = markLeft
            .Top = topPos + 1
            .Width = 14
            .Height = 16
            .Caption = ChrW(&HE8A5&)
            .Font.Name = "Segoe MDL2 Assets"
            .Font.Size = 10
            .ForeColor = nameColor
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        markLbl.ZOrder 0
        RegisterAttachClick markLbl
    End If

    Set nameLbl = list.Controls.Add( _
        "Forms.Label.1", _
        "att_" & itemId, _
        True)

    With nameLbl
        .Left = nameLeft
        .Top = topPos + 2
        .Height = 16
        .Caption = mAttachItems(index).Caption
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Font.Italic = (mAttachItems(index).Kind = "note")
        .ForeColor = nameColor
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .WordWrap = False
        .Width = rowWidth - nameLeft - 8
    End With

    nameLbl.ZOrder 0
    RegisterAttachClick nameLbl

    nameWidth = MeasureLabelWidth(nameLbl)
    nameLbl.AutoSize = False
    nameLbl.Height = 16
    nameLbl.Top = topPos + 2
    nameLbl.Width = nameWidth + 2
    cursorLeft = nameLeft + nameLbl.Width + 12

    If Len(mAttachItems(index).Status) > 0 Then
        Set statusLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "st_" & itemId, _
            True)

        With statusLbl
            .Top = topPos + 3
            .Height = 16
            .Caption = mAttachItems(index).Status
            .Font.Name = "Meiryo UI"
            .Font.Size = 8
            .ForeColor = statusColor
            .TextAlign = fmTextAlignLeft
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        statusWidth = MeasureLabelWidth(statusLbl)
        statusLbl.AutoSize = False
        statusLbl.Height = 16
        statusLbl.Top = topPos + 3
        statusLbl.Width = statusWidth + 2
        statusLbl.Left = cursorLeft
        cursorLeft = statusLbl.Left + statusLbl.Width + 10
        statusLbl.ZOrder 0
        RegisterAttachClick statusLbl
    End If

    If Len(PasswordCaption(mAttachItems(index).PasswordState)) > 0 Then
        Set pwLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "pw_" & itemId, _
            True)

        With pwLbl
            .Top = topPos + 3
            .Height = 16
            .Caption = PasswordCaption(mAttachItems(index).PasswordState)
            .Font.Name = "Meiryo UI"
            .Font.Size = 8
            .ForeColor = statusColor
            .TextAlign = fmTextAlignLeft
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        statusWidth = MeasureLabelWidth(pwLbl)
        pwLbl.AutoSize = False
        pwLbl.Height = 16
        pwLbl.Top = topPos + 3
        pwLbl.Width = statusWidth + 2
        pwLbl.Left = cursorLeft
        cursorLeft = pwLbl.Left + pwLbl.Width
        pwLbl.ZOrder 0
        RegisterAttachClick pwLbl
    End If

    tailRight = cursorLeft + 8
    If tailRight > rowWidth Then
        nameLbl.Width = nameLbl.Width - (tailRight - rowWidth)
        If nameLbl.Width < 24 Then nameLbl.Width = 24
        cursorLeft = nameLeft + nameLbl.Width + 12
        If Not statusLbl Is Nothing Then
            statusLbl.Left = cursorLeft
            cursorLeft = statusLbl.Left + statusLbl.Width + 10
        End If
        If Not pwLbl Is Nothing Then
            pwLbl.Left = cursorLeft
        End If
    End If

End Sub


Private Function PasswordCaption(ByVal passwordState As String) As String

    Select Case passwordState
        Case "set"
            PasswordCaption = "パスワード設定あり"
        Case "none"
            PasswordCaption = "パスワード設定なし"
    End Select

End Function


Private Sub RegisterAttachClick(ByVal ctl As Object)

    Dim handler As CCardClick

    Set handler = New CCardClick
    Set handler.ParentForm = Me

    If TypeName(ctl) = "Frame" Then
        Set handler.CardFrame = ctl
    Else
        Set handler.CardLabel = ctl
    End If

    mAttachHandlers.Add handler

End Sub


Private Sub ResizeRecipientColumns(ByVal attachTop As Single)

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim areaHeight As Single

    names = Array("fraExternal", "fraOutside", "fraInternal")
    areaHeight = attachTop - GAP - 30
    If areaHeight < 80 Then areaHeight = 80

    For i = LBound(names) To UBound(names)
        Set fra = Me.Controls(CStr(names(i)))
        fra.Height = areaHeight
        fra.Controls(fra.Name & "_list").Height = areaHeight - 34 - 1
        SyncScrollThumb fra.Controls(fra.Name & "_list")
    Next i

End Sub


Private Sub ClearFrameControls(ByVal fra As MSForms.Frame)

    Do While fra.Controls.Count > 0
        fra.Controls.Remove 0
    Loop

End Sub


'========================================================
' 細いスクロールバー
'========================================================
Public Sub SyncScrollThumb(ByVal list As Object)

    Dim host As Object
    Dim track As MSForms.Label
    Dim thumb As MSForms.Label
    Dim view As Single
    Dim maxTop As Single
    Dim thumbH As Single
    Dim travel As Single

    If list Is Nothing Then Exit Sub

    Set host = list.Parent
    Set track = FindHostControl(host, "bar_" & list.Name)
    Set thumb = FindHostControl(host, "thumb_" & list.Name)

    If track Is Nothing Or thumb Is Nothing Then
        Set track = host.Controls.Add( _
            "Forms.Label.1", _
            "bar_" & list.Name, _
            True)

        With track
            .Caption = ""
            .BackColor = RGB(214, 216, 220)
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        Set thumb = host.Controls.Add( _
            "Forms.Label.1", _
            "thumb_" & list.Name, _
            True)

        With thumb
            .Caption = ""
            .BackColor = RGB(132, 134, 140)
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        BindScrollDrag CStr(list.Name), track, thumb
    End If

    view = list.InsideHeight
    If view < 20 Then view = list.Height

    track.Left = list.Left + list.Width + 2
    track.Top = list.Top + 4
    track.Width = 6
    track.BackColor = RGB(214, 216, 220)
    If list.Height > 16 Then
        track.Height = list.Height - 12
    Else
        track.Height = list.Height
    End If

    thumb.Left = track.Left + 1
    thumb.Width = 4
    thumb.BackColor = RGB(132, 134, 140)

    maxTop = list.ScrollHeight - view

    If maxTop <= 1 Then
        track.Visible = False
        thumb.Visible = False
        Exit Sub
    End If

    track.Visible = True
    thumb.Visible = True

    thumbH = track.Height * (view / list.ScrollHeight)
    If thumbH < 28 Then thumbH = 28
    If thumbH > track.Height Then thumbH = track.Height
    thumb.Height = thumbH

    travel = track.Height - thumbH
    If list.ScrollTop < 0 Then list.ScrollTop = 0
    If list.ScrollTop > maxTop Then list.ScrollTop = maxTop

    If travel < 1 Then
        thumb.Top = track.Top
    Else
        thumb.Top = track.Top + (list.ScrollTop / maxTop) * travel
    End If

    track.ZOrder 0
    thumb.ZOrder 0

End Sub


Public Sub MoveScrollThumb( _
    ByVal listName As String, _
    ByVal delta As Single)

    Dim list As Object
    Dim track As MSForms.Label
    Dim thumb As MSForms.Label
    Dim view As Single
    Dim maxTop As Single
    Dim travel As Single
    Dim newTop As Single

    Set list = ScrollListByName(listName)
    If list Is Nothing Then Exit Sub

    Set track = list.Parent.Controls("bar_" & listName)
    Set thumb = list.Parent.Controls("thumb_" & listName)

    view = list.InsideHeight
    If view < 20 Then view = list.Height
    maxTop = list.ScrollHeight - view
    travel = track.Height - thumb.Height
    If maxTop <= 1 Or travel < 1 Then Exit Sub

    newTop = thumb.Top + delta
    If newTop < track.Top Then newTop = track.Top
    If newTop > track.Top + travel Then newTop = track.Top + travel

    thumb.Top = newTop
    list.ScrollTop = (newTop - track.Top) / travel * maxTop

End Sub


Public Sub JumpScrollThumb( _
    ByVal listName As String, _
    ByVal y As Single)

    Dim list As Object
    Dim track As MSForms.Label
    Dim thumb As MSForms.Label
    Dim delta As Single

    Set list = ScrollListByName(listName)
    If list Is Nothing Then Exit Sub

    Set track = list.Parent.Controls("bar_" & listName)
    Set thumb = list.Parent.Controls("thumb_" & listName)

    delta = (track.Top + y - thumb.Height / 2) - thumb.Top
    MoveScrollThumb listName, delta

End Sub


Private Sub BindScrollDrag( _
    ByVal listName As String, _
    ByVal track As MSForms.Label, _
    ByVal thumb As MSForms.Label)

    Dim handler As CScrollDrag
    Dim existing As Object

    On Error Resume Next
    Set existing = mScrollDrags.Item(listName)
    On Error GoTo 0
    If Not existing Is Nothing Then Exit Sub

    Set handler = New CScrollDrag
    Set handler.Track = track
    Set handler.Thumb = thumb
    Set handler.ParentForm = Me
    handler.ListName = listName
    mScrollDrags.Add handler, listName

End Sub


Private Function ScrollListByName(ByVal listName As String) As Object

    Dim hostName As String

    If Len(listName) <= 5 Then Exit Function
    hostName = Left$(listName, Len(listName) - 5)

    On Error Resume Next
    Set ScrollListByName = Me.Controls(hostName).Controls(listName)

End Function


Private Function FindHostControl( _
    ByVal host As Object, _
    ByVal controlName As String) As Object

    On Error Resume Next
    Set FindHostControl = host.Controls(controlName)

End Function


Private Function AttachItemVisible(ByVal index As Long) As Boolean

    Dim parentIndex As Long

    If Len(mAttachItems(index).ParentId) = 0 Then
        AttachItemVisible = True
        Exit Function
    End If

    parentIndex = FindAttachItem(mAttachItems(index).ParentId)
    If parentIndex = 0 Then Exit Function

    AttachItemVisible = AttachItemVisible(parentIndex) And AttachItemIsExpanded(mAttachItems(index).ParentId)

End Function


Private Function AttachItemIsExpanded(ByVal itemId As String) As Boolean

    Dim saved As String

    On Error Resume Next
    saved = CStr(mCollapsedIds(itemId))
    If Err.Number <> 0 Then
        AttachItemIsExpanded = True
        Err.Clear
    End If

End Function


Private Sub ExpandAllAttachments()

    Set mCollapsedIds = New Collection

End Sub


Private Sub CollapseAttachment(ByVal attachId As String)

    On Error Resume Next
    mCollapsedIds.Remove attachId
    On Error GoTo 0
    mCollapsedIds.Add attachId, attachId

End Sub


Private Sub ExpandAttachment(ByVal attachId As String)

    On Error Resume Next
    mCollapsedIds.Remove attachId
    On Error GoTo 0

End Sub


Private Function FindAttachItem(ByVal itemId As String) As Long

    Dim i As Long

    For i = 1 To mAttachCount
        If mAttachItems(i).Id = itemId Then
            FindAttachItem = i
            Exit Function
        End If
    Next i

End Function


Private Sub AttachRowStyle( _
    ByVal index As Long, _
    ByRef nameColor As Long, _
    ByRef statusColor As Long, _
    ByRef backColor As Long)

    nameColor = RGB(55, 55, 60)
    statusColor = RGB(55, 55, 60)
    backColor = RGB(255, 255, 255)

    ' 安全なのは、パスワード付き ZIP の中にある PDF だけ
    If AttachmentIsSafe(index) Then
        nameColor = RGB(70, 120, 78)
        statusColor = RGB(70, 120, 78)
    Else
        backColor = RGB(252, 244, 243)
    End If

End Sub


'========================================================
' フッター
'========================================================
Private Sub CreateFooter()

    Dim topPos As Single
    Dim sendLeft As Single

    topPos = Me.InsideHeight - MARGIN - 28
    sendLeft = Me.InsideWidth - MARGIN - 96

    '表示の初期化。枠だけの静かなボタンにして、送信と並べない
    CreateFlatButton _
        "btnReset", _
        "表示を初期化", _
        ChrW(&H21BA), _
        MARGIN, _
        topPos, _
        132, _
        28, _
        RGB(255, 255, 255), _
        RGB(90, 92, 98), _
        RGB(214, 216, 220), _
        RGB(236, 237, 240), _
        False, _
        "Reset"
    Me.Controls("btnReset").ControlTipText = _
        "列の絞り込み、開いているカード、ZIP の展開、スクロールを最初の状態に戻します"

    CreateFlatButton _
        "btnCancel", _
        "キャンセル", _
        "", _
        sendLeft - 8 - 96, _
        topPos, _
        96, _
        28, _
        RGB(255, 255, 255), _
        RGB(45, 45, 48), _
        RGB(214, 216, 220), _
        RGB(236, 237, 240), _
        False, _
        "Cancel"

    CreateFlatButton _
        "btnSend", _
        "送信", _
        "", _
        sendLeft, _
        topPos, _
        96, _
        28, _
        RGB(32, 33, 36), _
        RGB(255, 255, 255), _
        RGB(32, 33, 36), _
        RGB(16, 16, 18), _
        True, _
        "Send"

End Sub


'========================================================
' フラットなフッターボタン
'========================================================
Private Sub CreateFlatButton( _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal iconChar As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal width As Single, _
    ByVal height As Single, _
    ByVal backColor As Long, _
    ByVal foreColor As Long, _
    ByVal borderColor As Long, _
    ByVal downColor As Long, _
    ByVal bold As Boolean, _
    ByVal actionName As String)

    Dim plate As MSForms.Frame
    Dim face As MSForms.Label
    Dim icon As MSForms.Label

    Set plate = Me.Controls.Add( _
        "Forms.Frame.1", _
        controlName, _
        True)

    With plate
        .Left = leftPos
        .Top = topPos
        .Width = width
        .Height = height

        .Caption = ""

        .BackColor = backColor
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = borderColor
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
    End With

    If Len(iconChar) > 0 Then
        Set icon = plate.Controls.Add( _
            "Forms.Label.1", _
            controlName & "_icon", _
            True)

        With icon
            .Caption = iconChar
            .AutoSize = False
            .WordWrap = False
            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
            .ForeColor = foreColor
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = 11
        End With
    End If

    Set face = plate.Controls.Add( _
        "Forms.Label.1", _
        controlName & "_face", _
        True)

    With face
        .Caption = caption
        .AutoSize = False
        .WordWrap = False
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
        .ForeColor = foreColor
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Font.Bold = bold
    End With

    CenterButtonCaption plate, face, icon

    If Not icon Is Nothing Then
        RegisterFlatClick icon, plate, backColor, downColor, actionName
    End If
    RegisterFlatClick plate, plate, backColor, downColor, actionName
    RegisterFlatClick face, plate, backColor, downColor, actionName

End Sub


' 文字の実寸を測って、ボタンの内側の中央に置く
Private Sub CenterButtonCaption( _
    ByVal plate As MSForms.Frame, _
    ByVal face As MSForms.Label, _
    ByVal icon As MSForms.Label)

    Dim gap As Single
    Dim groupW As Single
    Dim groupLeft As Single
    Dim textW As Single
    Dim iconW As Single
    Dim boxW As Single
    Dim boxH As Single

    gap = 4
    boxW = plate.Width - 2
    boxH = plate.Height - 2

    textW = MeasureLabelWidth(face) + 2
    face.TextAlign = fmTextAlignCenter
    face.Width = textW
    face.Height = face.Font.Size + 6
    face.Top = (boxH - face.Font.Size) / 2

    If icon Is Nothing Then
        face.Left = (boxW - textW) / 2
        Exit Sub
    End If

    iconW = MeasureLabelWidth(icon) + 2
    icon.TextAlign = fmTextAlignCenter
    icon.Width = iconW
    icon.Height = icon.Font.Size + 6
    icon.Top = (boxH - icon.Font.Size) / 2

    groupW = iconW + gap + textW
    groupLeft = (boxW - groupW) / 2
    If groupLeft < 2 Then groupLeft = 2

    icon.Left = groupLeft
    face.Left = groupLeft + iconW + gap

End Sub


Private Function MeasureLabelWidth(ByVal lbl As MSForms.Label) As Single

    On Error GoTo Fallback

    lbl.WordWrap = False
    lbl.AutoSize = False
    lbl.Width = 200
    lbl.Height = 24
    lbl.AutoSize = True
    MeasureLabelWidth = lbl.Width
    lbl.AutoSize = False

    If MeasureLabelWidth < 8 Then MeasureLabelWidth = 8
    Exit Function

Fallback:
    MeasureLabelWidth = lbl.Font.Size * Len(lbl.Caption) + 4
    lbl.AutoSize = False

End Function


'========================================================
' Buttonイベント登録
'========================================================
Private Sub RegisterFlatClick( _
    ByVal ctl As Object, _
    ByVal host As Object, _
    ByVal idleBack As Long, _
    ByVal downBack As Long, _
    ByVal actionName As String)

    Dim handler As CButtonHandler

    Set handler = New CButtonHandler

    Set handler.ParentForm = Me
    Set handler.Host = host
    handler.ActionName = actionName
    handler.IdleBack = idleBack
    handler.DownBack = downBack

    If TypeName(ctl) = "Frame" Then
        Set handler.Plate = ctl
    Else
        Set handler.Face = ctl
    End If

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

        Case "Reset"

            ResetDisplay

    End Select

End Sub


'========================================================
' 画面の開き方だけを、起動直後の状態に戻す
'========================================================
Private Sub ResetDisplay()

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame

    If Not mExpandedCard Is Nothing Then
        ApplyCardLayout mExpandedCard, False
        Set mExpandedCard = Nothing
    End If

    mCategoryFilter = ""
    mAttachExpanded = False
    ApplyCategoryFilter

    names = Array("fraExternal", "fraOutside", "fraInternal")

    For i = LBound(names) To UBound(names)
        Set fra = Me.Controls(CStr(names(i)))
        Set list = fra.Controls(fra.Name & "_list")
        list.ScrollTop = 0
        SyncScrollThumb list
    Next i

    ExpandAllAttachments
    RenderAttachmentList

    Set list = Me.Controls("fraAttachments").Controls("fraAttachments_list")
    list.ScrollTop = 0

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
