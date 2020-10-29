Imports System
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Visa

    ''' <summary>
    ''' Управление измерительными приборами по спецификации VISA (Virtual Instruments Software Architecture).
    ''' </summary>
    Public Class VisaDevice
        Implements IDisposable

        ''' <summary>
        ''' Выбранный тип VISA. От него зависит, к какой библиотеке будет обращаться программа.
        ''' </summary>
        Private Shared M_defRM As ManufacturerIds = ManufacturerIds.None

        ''' <summary>
        ''' Идентификатор сессии с заданным ресурсом. Определяется после открытия соединения в <see cref="Native.viOpen(Integer, String, Native.AccessMode, Integer, ByRef Integer)"/>.
        ''' </summary>
        Private Session As Integer = Native.VI_ERROR_INV_SESSION

#Region "PROPS"

        ''' <summary>
        ''' Уникальный идентификатор ресурса.
        ''' </summary>
        Public ReadOnly Property Description As String
            Get
                Return _Description
            End Get
        End Property
        Private _Description As String

        Public ReadOnly Property Manufacturer As String
            Get
                Return _Manufacturer
            End Get
        End Property
        Private _Manufacturer As String

        ''' <summary>
        ''' Тип интерфейса.
        ''' </summary>
        Public ReadOnly Property IntfType As InterfaceType
            Get
                Return _InterfaceType
            End Get
        End Property
        Private _InterfaceType As InterfaceType = VisaDevice.InterfaceType.None

        ''' <summary>
        ''' Класс ресурса.
        ''' </summary>
        Public ReadOnly Property ResourceClass As String
            Get
                Return _ResourceClass
            End Get
        End Property
        Private _ResourceClass As String

#End Region '/PROPS

#Region "CTOR"

        Shared Sub New()
            CheckLibraries()
            ChangeDefaultLibrary(ManufacturerIds.MANFID_DEFAULT)
        End Sub

        ''' <summary>
        ''' Подключается к заданному ресурсу.
        ''' </summary>
        ''' <param name="name">Уникальное символьное имя устройства.</param>
        Public Sub New(name As String)
            _Description = name
            OpenDevice(name)
            Dim buf As New StringBuilder(Native.VI_FIND_BUFLEN)
            Dim status As ViStatus = Native.viGetAttribute(Session, ViAttr.VI_ATTR_MANF_NAME, buf)
            _Manufacturer = buf.ToString()
        End Sub

#End Region '/CTOR

#Region "SHARED METHODS"

        ''' <summary>
        ''' Установленные библиотеки для устройств VISA, которые можно использовать для обмена.
        ''' </summary>
        Public Shared Function GetVisaLibraries() As String()
            If (_VisaLibraries.Count = 0) Then
                CheckLibraries()
            End If
            Return _VisaLibraries.ToArray()
        End Function
        Private Shared _VisaLibraries As New List(Of String)

        ''' <summary>
        ''' Найденные устройства.
        ''' </summary>
        Public Shared Function GetVisaDevices() As String()
            If (_VisaDevices.Count = 0) Then
                FindDevices()
            End If
            Return _VisaDevices.ToArray()
        End Function
        Private Shared _VisaDevices As New List(Of String)

        ''' <summary>
        ''' Выгружает из памяти библиотеку. Вызывается после завершения работы со всеми устройствами VISA.
        ''' </summary>
        Public Shared Sub UnloadLibrary()
            Native.RsViUnloadVisaLibrary()
        End Sub

        ''' <summary>
        ''' Меняет библиотеку по умолчанию на заданную <paramref name="useLib"/> (при необходимости). По умолчанию используется стандартная библиотека.
        ''' </summary>
        Public Shared Sub ChangeDefaultLibrary(useLib As ManufacturerIds)
            Dim res As ViStatus
            res = Native.viClose(M_defRM)
            Native.RsViUnloadVisaLibrary()
            Dim resDef = Native.RsViSetDefaultLibrary(useLib)
            Dim rmTmp As Integer
            res = Native.viOpenDefaultRM(rmTmp)
            M_defRM = CType(rmTmp, ManufacturerIds)
        End Sub

        ''' <summary>
        ''' Проверяет, какие библиотеки установлены в системе.
        ''' </summary>
        Private Shared Sub CheckLibraries()
            _VisaLibraries = New List(Of String)
            If IsVisaLibraryInstalled(ManufacturerIds.MANFID_DEFAULT) Then
                _VisaLibraries.Add("Default")
            End If
            If IsVisaLibraryInstalled(ManufacturerIds.MANFID_NI) Then
                _VisaLibraries.Add("National Instruments")
            End If
            If IsVisaLibraryInstalled(ManufacturerIds.MANFID_RS) Then
                _VisaLibraries.Add("Rohde&Schwarz")
            End If
            If IsVisaLibraryInstalled(ManufacturerIds.MANFID_AG) Then
                _VisaLibraries.Add("Keysight")
            End If
        End Sub

        ''' <summary>
        ''' Ищет поддерживаемые устройства.
        ''' </summary>
        Private Shared Sub FindDevices()
            _VisaDevices = New List(Of String)
            Dim attr As RsFindModes = RsFindModes.VI_RS_FIND_MODE_CONFIG Or RsFindModes.VI_RS_FIND_MODE_MDNS Or RsFindModes.VI_RS_FIND_MODE_VXI11
            Dim res As ViStatus = Native.viSetAttribute(M_defRM, ViAttr.VI_RS_ATTR_TCPIP_FIND_RSRC_MODE, attr)

            Dim findList As Integer = 0
            Dim retCount As Integer = 0
            Dim desc As New StringBuilder(Native.VI_FIND_BUFLEN)
            res = Native.viFindRsrc(M_defRM, "?*", findList, retCount, desc)
            If (retCount > 0) Then
                _VisaDevices.Add(desc.ToString())
                For i As Integer = 0 To retCount - 1 - 1
                    Native.viFindNext(findList, desc)
                    _VisaDevices.Add(desc.ToString())
                Next
            End If
        End Sub

        ''' <summary>
        ''' Проверяет, имеется ли в системе библиотека, заданная <paramref name="manId"/>.
        ''' </summary>
        Private Shared Function IsVisaLibraryInstalled(manId As ManufacturerIds) As Boolean
            Return (Native.RsViIsVisaLibraryInstalled(manId) = Native.VI_BOOL.VI_TRUE)
        End Function

#End Region '/SHARED METHODS

#Region "OPEN METHODS"

        ''' <summary>
        ''' Запрашивает идентификатор IDN у текущего устройства.
        ''' </summary>
        Public Function ShowIdn() As String
            Dim answer As String = SendQuery("*IDN?")
            Return answer
        End Function

        ''' <summary>
        ''' Отправляет запрос и возвращает ответ.
        ''' </summary>
        ''' <param name="query">Запрос.</param>
        Public Function SendQuery(query As String) As String
            Write(query & vbLf)
            If query.Contains("?") Then
                Dim answer As String = Read()
                Return answer
            End If
            Return ""
        End Function

#End Region '/OPEN METHODS

#Region "CLOSED METHODS"

        ''' <summary>
        ''' Открывает заданное устройство и назначает номер сессии.
        ''' </summary>
        ''' <param name="name">Уникальное символьное имя устройства.</param>
        Private Sub OpenDevice(name As String)
            Dim status As ViStatus = Native.viOpen(M_defRM, name, Native.AccessMode.VI_NO_LOCK, 0, Session)
            CheckResult(status)

            status = Native.viSetAttribute(Session, ViAttr.VI_ATTR_SEND_END_EN, Native.VI_BOOL.VI_TRUE)
            CheckResult(status)

            Try
                Dim intfNum As Short
                Dim rsrcClass As New StringBuilder(Native.VI_FIND_BUFLEN)
                Dim unaliasName As New StringBuilder(Native.VI_FIND_BUFLEN)
                Dim aliasName As New StringBuilder(Native.VI_FIND_BUFLEN)
                Dim res As ViStatus = Native.viParseRsrcEx(M_defRM, name, _InterfaceType, intfNum, rsrcClass, unaliasName, aliasName)
                Debug.WriteLine($"intf ty={IntfType}, intf num={intfNum}, class={rsrcClass}, unaliased name={unaliasName}, alias={aliasName}")
                _ResourceClass = rsrcClass.ToString()
            Catch ex As Exception
                Dim intfNum As Short
                Dim res As ViStatus = Native.viParseRsrc(M_defRM, name, _InterfaceType, intfNum)
                Debug.WriteLine($"intf ty={IntfType}, intf num={intfNum}")
            End Try
        End Sub

        ''' <summary>
        ''' Передаёт строку <paramref name="query"/> текущему устройству.
        ''' </summary>
        ''' <param name="query"></param>
        Private Sub Write(query As String)
            Dim retCount As Integer
            CheckResult(Native.viWrite(Session, query, query.Length, retCount))
        End Sub

        ''' <summary>
        ''' Читает ответ заданного устройства.
        ''' </summary>
        Private Function Read() As String
            Dim status As ViStatus
            Dim retSb As New StringBuilder(1024)
            Do
                Dim retCount As Integer
                Dim sb As New StringBuilder(1024)
                status = Native.viRead(Session, sb, sb.Capacity, retCount)
                Debug.WriteLine(status)
                If (retCount > 0) Then
                    retSb.Append(sb.ToString(0, retCount))
                End If
            Loop While (status = ViStatus.VI_SUCCESS_MAX_CNT)
            Return retSb.ToString()
        End Function


#End Region '/CLOSED METHODS

#Region "HELPERS"

        ''' <summary>
        ''' Проверяет статус.
        ''' </summary>
        ''' <param name="status">Код ошибки/результат выполнения функции.</param>
        ''' <remarks>
        ''' Код ошибки:
        ''' 0 - без ошибок;
        ''' &lt; 0 - ошибки;
        ''' > 0 - предупреждения.
        ''' </remarks>
        Private Sub CheckResult(status As ViStatus)
            Select Case status
                Case ViStatus.VI_SUCCESS
                    Return

                Case Is < ViStatus.VI_SUCCESS
                    Dim message As New StringBuilder(Native.VI_FIND_BUFLEN)
                    Dim res As ViStatus = Native.viStatusDesc(M_defRM, status, message)
                    Throw New Exception(message.ToString())

                Case Is > ViStatus.VI_SUCCESS
                    Dim message As New StringBuilder(Native.VI_FIND_BUFLEN)
                    Dim res As ViStatus = Native.viStatusDesc(M_defRM, status, message)
                    Debug.WriteLine("ЗАМЕЧАНИЕ: " & message.ToString())

            End Select
        End Sub

#End Region '/HELPERS

#Region "IDISPOSABLE SUPPORT"

        Private DisposedValue As Boolean

        Private Sub Dispose(disposing As Boolean)
            If (Not DisposedValue) Then
                If disposing Then
                    
                End If
                CheckResult(Native.viClose(Session)) 
                DisposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region '/IDISPOSABLE SUPPORT

#Region "NESTED TYPES"

        ''' <summary>
        ''' Нативные функции.
        ''' </summary>
        Private NotInheritable Class Native

            Private Const LibPath As String = "RsVisaLoader.dll"

#Region "РАБОТА С БИБЛИОТЕКОЙ"

            ''' <summary>
            ''' Проверяет установлена ли библиотека VISA для производителя <paramref name="manfId"/>. 
            ''' Возвращает <see cref="VI_BOOL.VI_TRUE"/>, если библиотека установлена, и <see cref="VI_BOOL.VI_FALSE"/> в противном случае.
            ''' </summary>
            ''' <param name="manfId">Идентификатор производителя.</param>
            <DllImport(LibPath, EntryPoint:="#062", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function RsViIsVisaLibraryInstalled(manfId As ManufacturerIds) As VI_BOOL
            End Function

            ''' <summary>
            ''' Выбирает библиотеку VISA по умолчанию.
            ''' Возвращает <see cref="Native.VI_BOOL.VI_TRUE"/>, если библиотека была изменена успешно, и <see cref="VI_BOOL.VI_FALSE"/> в противном случае.
            ''' </summary>
            ''' <param name="manfId">A value Of RSVISA_MANFID_DEFAULT, RSVISA_MANFID_RS, etc.</param>
            ''' <remarks>
            ''' Необходимо вызвать этот метод прежде чем обращаться к другим функциям библиотеки VISA (в т.ч. <see cref="Native.viOpenDefaultRM(ByRef Integer)"/>).
            ''' Если библиотека по умолчанию не может быть загружена, производится попытка использовать rsvisa32.dll.
            ''' Библиотека загружается при первом обращении к функциям VISA.
            ''' </remarks>
            <DllImport(LibPath, EntryPoint:="#060", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function RsViSetDefaultLibrary(manfId As ManufacturerIds) As VI_BOOL
            End Function

            ''' <summary>
            ''' Выгружает библиотеку VISA. 
            ''' </summary>
            ''' <remarks>
            ''' Этот метод вызывается только после вызова <see cref="viClose(Integer)"/>.
            ''' </remarks>
            <DllImport(LibPath, EntryPoint:="#061", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Sub RsViUnloadVisaLibrary()
            End Sub

#End Region '/РАБОТА С БИБЛИОТЕКОЙ

#Region "РАБОТА С РЕСУРСАМИ"

            ''' <summary>
            ''' Возвращает сессию ресурсу менеджера ресурсов (RM = resource manager) по умолчанию.
            ''' </summary>
            ''' <param name="sesn">Уникальный логический идентификатор сессии менеджера ресурсов по умолчанию.</param>
            ''' <returns></returns>
            ''' <remarks>
            ''' Этф ф-ция должна быть вызвана первой. Первый вызов этой ф-ции инициализирует систему VISA.
            ''' </remarks>
            <DllImport(LibPath, EntryPoint:="#141", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viOpenDefaultRM(<Out()> ByRef sesn As Integer) As ViStatus
            End Function

            <DllImport(LibPath, EntryPoint:="#128", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGetDefaultRM(<Out()> ByRef sesn As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Запрашивает у системы VISA найти ресурсы, связанные со специфическим идентификатором.
            ''' </summary>
            ''' <param name="sesn">Уникальный логический идентификатор сессии менеджера ресурсов по умолчанию.</param>
            ''' <param name="expr">Регулярное выражение, по которому будет идти поиск ресурсов. Примеры: "?*", "VXI?*INSTR", "?*INSTR", "VXI0::?*", "(GPIB|VXI)?*INSTR", "ASRL[0-9]+::INSTR".</param>
            ''' <param name="findList">Возвращает указатель, идентифицирующий текущую сессию поиска. Этот указатель используется в <see cref="viFindNext(Integer, StringBuilder)"/>.</param>
            ''' <param name="retCount">Число совпадений.</param>
            ''' <param name="descr">Строка, идентифицирующая расположение ресурса. Минимальная длина <see cref="VI_FIND_BUFLEN"/> байтов.
            ''' Это описание используется в <see cref="viOpen(Integer, String, AccessMode, Integer, ByRef Integer)"/>.</param>
            <DllImport(LibPath, EntryPoint:="#129", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viFindRsrc(sesn As Integer, expr As String, ByRef findList As Integer, ByRef retCount As Integer, descr As StringBuilder) As ViStatus
            End Function

            ''' <summary>
            ''' Возвращает следующий ресурс, начиная с указателя <paramref name="findList"/>, который был определён в предыдущем вызове <see cref="viFindRsrc(Integer, String, ByRef Integer, ByRef Integer, StringBuilder)"/>.
            ''' </summary>
            ''' <param name="findList">Указатель на список ресурсов. Получается из <see cref="viFindRsrc(Integer, String, ByRef Integer, ByRef Integer, StringBuilder)"/>.</param>
            ''' <param name="descr">Строка, идентифицирующая ресурс. Минимальная длина <see cref="VI_FIND_BUFLEN"/> байтов. 
            ''' Передаётся в <see cref="viOpen(Integer, String, AccessMode, Integer, ByRef Integer)"/>, чтобы начать сессию с данным ресурсом.</param>
            <DllImport(LibPath, EntryPoint:="#130", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viFindNext(findList As Integer, descr As StringBuilder) As ViStatus
            End Function

            ''' <summary>
            ''' Разбирает строку, идентифицирующую ресурс, чтобы получить информацию об интерфейсе.
            ''' </summary>
            ''' <param name="sesn">Сессия менеджера ресурсов.</param>
            ''' <param name="rsrcName">Уникальное символьное имя ресурса.</param>
            ''' <param name="intfType">Тип интерфейса для данного ресурса.</param>
            ''' <param name="intfNum">Номер интерфейсной платы данного ресурса.</param>
            <DllImport(LibPath, EntryPoint:="#146", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viParseRsrc(sesn As Integer, rsrcName As String, ByRef intfType As InterfaceType, ByRef intfNum As Short) As ViStatus
            End Function

            ''' <summary>
            ''' Разбирает строку, идентифицирующую ресурс, чтобы получить расширенную информацию об интерфейсе.
            ''' </summary>
            ''' <param name="sesn">Сессия менеджера ресурсов.</param>
            ''' <param name="rsrcName">Уникальное символьное имя ресурса.</param>
            ''' <param name="intfType">Тип интерфейса для данного ресурса.</param>
            ''' <param name="intfNum">Номер интерфейсной платы данного ресурса.</param>
            ''' <param name="rsrcClass">Класс ресурса (например, "INSTR").</param>
            ''' <param name="expandedUnaliasedName">Расширенное имя ресурса.</param>
            ''' <param name="aliasIfExists">Задаёт пользовательский псевдоним ресурса (если есть).</param>
            <DllImport(LibPath, EntryPoint:="#147", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viParseRsrcEx(sesn As Integer, rsrcName As String, <Out()> ByRef intfType As InterfaceType, <Out()> ByRef intfNum As Short, rsrcClass As StringBuilder, expandedUnaliasedName As StringBuilder, aliasIfExists As StringBuilder) As ViStatus
            End Function

            ''' <summary>
            ''' Открывает сессию для заданного ресурса <paramref name="sesn"/>. 
            ''' Присвоенный идентификатор сессии <paramref name="vi"/> в дальнейшем будет применяться для вызова всех остальных операций с данным ресурсом.
            ''' </summary>
            ''' <param name="sesn">Сессия менеджера ресурсов, которую возвращает <see cref="viOpenDefaultRM(ByRef Integer)"/>.</param>
            ''' <param name="descr">Уникальное символьное имя ресурса.</param>
            ''' <param name="mode">Задаёт режим доступа к ресурсу.</param>
            ''' <param name="timeout">Время в мс, по истечению которого будет возвращена ошибка, если не будет получен ответ. Этот параметр не задаёт таймаут I/O.</param>
            ''' <param name="vi">Уникальный логический идентификатор сессии. Выходное значение.</param>
            <DllImport(LibPath, EntryPoint:="#131", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viOpen(sesn As Integer, descr As String, mode As AccessMode, timeout As Integer, <Out()> ByRef vi As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Закрывает сессию, событие или список, заданный идентификатором <paramref name="vi"/>. Освобождает все задействованные структуры данных.
            ''' </summary>
            ''' <param name="vi">Уникальный логический идентификатор сессии, события или списка.</param>
            <DllImport(LibPath, EntryPoint:="#132", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viClose(vi As Integer) As ViStatus
            End Function

#End Region '/РАБОТА С РЕСУРСАМИ

#Region "ПОЛУЧЕНИЕ И УСТАНОВКА АТРИБУТОВ"

            ''' <summary>
            ''' Получает значение заданного атрибута <paramref name="attrName"/>.
            ''' </summary>
            <DllImport(LibPath, EntryPoint:="#133", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGetAttribute(vi As Integer, attrName As ViAttr, <Out()> attrValue As Byte) As ViStatus
            End Function

            ''' <summary>
            ''' Получает значение заданного атрибута <paramref name="attrName"/>.
            ''' </summary>
            <DllImport(LibPath, EntryPoint:="#133", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGetAttribute(vi As Integer, attrName As ViAttr, <Out()> ByRef attrValue As Short) As ViStatus
            End Function

            ''' <summary>
            ''' Получает значение заданного атрибута <paramref name="attrName"/>.
            ''' </summary>
            <DllImport(LibPath, EntryPoint:="#133", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGetAttribute(vi As Integer, attrName As ViAttr, <Out()> ByRef attrValue As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Получает значение заданного атрибута <paramref name="attrName"/>.
            ''' </summary>
            ''' <param name="vi">Идентификатор сессии, события или списка устройств.</param>
            ''' <param name="attrName">Имя атрибута, значение которого необходимо получить.</param>
            ''' <param name="attrValue">Значение атрибута. Выходное значение.</param>
            ''' <returns></returns>
            <DllImport(LibPath, EntryPoint:="#133", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGetAttribute(vi As Integer, attrName As ViAttr, attrValue As StringBuilder) As ViStatus
            End Function

            ''' <summary>
            ''' Устанавливает значение заданного атрибута <paramref name="attrName"/>.
            ''' </summary>
            ''' <param name="vi">Идентификатор сессии, события или списка.</param>
            ''' <param name="attrName">Атрибут, который следует установить.</param>
            ''' <param name="attrValue">Значение атрибута.</param>
            ''' <returns></returns>
            <DllImport(LibPath, EntryPoint:="#134", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viSetAttribute(vi As Integer, attrName As ViAttr, attrValue As Byte) As ViStatus
            End Function

            ''' <summary>
            ''' Устанавливает значение заданного атрибута <paramref name="attrName"/>.
            ''' </summary>
            <DllImport(LibPath, EntryPoint:="#134", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viSetAttribute(vi As Integer, attrName As ViAttr, attrValue As Short) As ViStatus
            End Function

            ''' <summary>
            ''' Устанавливает значение заданного атрибута <paramref name="attrName"/>.
            ''' </summary>
            <DllImport(LibPath, EntryPoint:="#134", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viSetAttribute(vi As Integer, attrName As ViAttr, attrValue As Integer) As ViStatus
            End Function

#End Region '/ПОЛУЧЕНИЕ И УСТАНОВКА АТРИБУТОВ

#Region "HELPERS"

            ''' <summary>
            ''' Возвращает человекочитаемое описание статуса.
            ''' </summary>
            ''' <param name="vi">Уникальный логический идентификатор сессии, события или списка.</param>
            ''' <param name="status">Код статуса, который необходимо интерпретировать.</param>
            ''' <param name="descr">Человекочитаемая строка, описывающая кода статуса переданной функции. Минимальная длина <see cref="VI_FIND_BUFLEN"/> байтов.</param>
            <DllImport(LibPath, EntryPoint:="#142", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viStatusDesc(vi As Integer, status As ViStatus, descr As StringBuilder) As ViStatus
            End Function

#End Region '/HELPERS

#Region "УПРАВЛЕНИЕ"

            ''' <summary>
            ''' Запрашивает сессию VISA прервать нормальное выполнение операции <paramref name="jobId"/>.
            ''' </summary>
            ''' <param name="vi">Идентификатор объекта.</param>
            ''' <param name="degree"><see cref="Native.VI_BOOL.VI_NULL"/>.</param>
            ''' <param name="jobId">Уникальный идентификатор операции, сгенерированный при вызове асинхронной операции.</param>
            <DllImport(LibPath, EntryPoint:="#143", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viTerminate(vi As Integer, degree As Short, jobId As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Блокирует ресурс <paramref name="vi"/>.
            ''' </summary>
            ''' <param name="vi">Идентификатор сессии.</param>
            ''' <param name="lockType">Тип блокировки - эксклюзивная или разделяемая.</param>
            ''' <param name="timeout">Время в мс или <see cref="Timeout"/> до возврата с ошибкой, если блокировка не получена.</param>
            ''' <param name="requestedKey">При эксклюзивной блокировке должно быть <see cref="Native.VI_BOOL.VI_NULL"/>. 
            ''' При разделяемой блокировке сессия может установить его в VI_NULL, тогда VISA генерирует <paramref name="accessKey"/> для сессии. 
            ''' Или сессия предложит <paramref name="accessKey"/> для использования разделяемой блокировки.</param>
            ''' <param name="accessKey">Выходной параметр. При разделяемой блокировке это значение используется для передачи другим сессиям. 
            ''' При эксклюзивной блокировке должно быть <see cref="Native.VI_BOOL.VI_NULL"/>.</param>
            <DllImport(LibPath, EntryPoint:="#144", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viLock(vi As Integer, lockType As Native.AccessMode, timeout As Integer, requestedKey As String, accessKey As StringBuilder) As ViStatus
            End Function

            ''' <summary>
            ''' Разблокирует ранее заблокированный функцией <see cref="viLock(Integer, AccessMode, Integer, String, StringBuilder)"/> ресурс.
            ''' </summary>
            ''' <param name="vi">Уникальный идентификатор сессии.</param>
            <DllImport(LibPath, EntryPoint:="#145", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viUnlock(vi As Integer) As ViStatus
            End Function

#End Region '/УПРАВЛЕНИЕ

#Region "ОБРАБОТКА СОБЫТИЙ"

            ''' <summary>
            ''' Прототип функции, которую вы хотите определить.
            ''' </summary>
            ''' <param name="vi">Уникальный логический идентификатор сессии.</param>
            ''' <param name="eventType">Логический идентификатор событияю</param>
            ''' <param name="context">Указатель, задающий уникальное событие.</param>
            ''' <param name="userHandle">Значение, заданное приложением, которое может использоваться для уникальной идентификации указателей события в сессии.</param>
            Public Delegate Function ViEventHandler(vi As Integer, eventType As ViEventType, context As Integer, userHandle As Integer) As ViStatus

            ''' <summary>
            ''' Активирует уведомление о событии типа <paramref name="eventType"/> для механизмов <paramref name="mechanism"/>.
            ''' </summary>
            ''' <param name="vi">Идентификатор сессии.</param>
            ''' <param name="eventType">Логический идентификатор события.</param>
            ''' <param name="mechanism">Задаёт механизмы обработки события. 
            ''' Механизм ожидания активируется с помощью <see cref="EventMechanism.VI_QUEUE"/>. 
            ''' Механизм обратного вызова - <see cref="EventMechanism.VI_HNDLR"/> или <see cref="EventMechanism.VI_SUSPEND_HNDLR"/>. 
            ''' Можно объединять по OR.</param>
            ''' <param name="context"><see cref="Native.VI_BOOL.VI_NULL"/>.</param>
            <DllImport(LibPath, EntryPoint:="#135", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viEnableEvent(vi As Integer, eventType As ViEventType, mechanism As EventMechanism, context As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Отменяет обслуживание события типа <paramref name="eventType"/> для механизма <paramref name="mechanism"/>.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="eventType">Логический идентификатор события.</param>
            ''' <param name="mechanism">Задаёт механизм, который должен быть отключён.</param>
            <DllImport(LibPath, EntryPoint:="#136", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viDisableEvent(vi As Integer, eventType As ViEventType, mechanism As EventMechanism) As ViStatus
            End Function

            ''' <summary>
            ''' Сбрасывает все ожидающие события типа <paramref name="eventType"/> для механизма <paramref name="mechanism"/>.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="eventType">Логический идентификатор события.</param>
            ''' <param name="mechanism">Задаёт механизм, для которого события должны быть сброшены. Можно использовать <see cref="EventMechanism.VI_ALL_MECH"/>.</param>
            <DllImport(LibPath, EntryPoint:="#137", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viDiscardEvents(vi As Integer, eventType As ViEventType, mechanism As EventMechanism) As ViStatus
            End Function

            ''' <summary>
            ''' Ожидает возникновение события в сессии <paramref name="vi"/>.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="eventType">Логический идентификатор события, которые следует ожидать.</param>
            ''' <param name="timeout">Таймаут ожидания события в мс.</param>
            ''' <param name="outEventType">Логический идентификатор полученного события.</param>
            ''' <param name="outEventContext">Уникальный указатель события.</param>
            <DllImport(LibPath, EntryPoint:="#138", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viWaitOnEvent(vi As Integer, eventType As ViEventType, timeout As Integer, <Out()> ByRef outEventType As ViEventType, <Out()> ByRef outEventContext As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Позволяет приложению установить обработчики событий.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="eventType">Логический идентификатор события.</param>
            ''' <param name="handler">Ссылка на обработчик.</param>
            ''' <param name="userHandle">Заданный приложением указатель, который может быть использован для идентификации обработчиков событий заданного типа.</param>
            <DllImport(LibPath, EntryPoint:="#139", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viInstallHandler(vi As Integer, eventType As ViEventType, handler As ViEventHandler, userHandle As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Позволяет приложению удалить обработчики событий.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="eventType">Логический идентификатор события.</param>
            ''' <param name="handler">Ссылка на обработчик, который необходимо удалить.</param>
            ''' <param name="userHandle">Заданный приложением указатель, который может быть использован для идентификации обработчиков событий заданного типа.</param>
            <DllImport(LibPath, EntryPoint:="#140", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viUninstallHandler(vi As Integer, eventType As ViEventType, handler As ViEventHandler, userHandle As Integer) As ViStatus
            End Function

#End Region '/ОБРАБОТКА СОБЫТИЙ

#Region "БАЗОВЫЕ ОПЕРАЦИИ ВВОДА-ВЫВОДА"

            ''' <summary>
            ''' Синхронно читает данные с устройства. Полученные данные хранятся в <paramref name="buffer"/>. Единовременно допустим запуск только одной операции чтения.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="buffer">Данные, полученные от устройства.</param>
            ''' <param name="bytesToRead">Число байтов, которые нужно прочитать.</param>
            ''' <param name="bytesWereRed">Число прочитанных байтов.</param>
            ''' <remarks>Чтение заканчивается, если: получен индикатор конца передачи; прочитан символ окончания передачи; прочитанное число данных равно <paramref name="bytesToRead"/>.</remarks>
            <DllImport(LibPath, EntryPoint:="#256", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viRead(vi As Integer, buffer As Byte(), bytesToRead As Integer, <Out()> ByRef bytesWereRed As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Синхронно читает данные с устройства. Полученные данные хранятся в <paramref name="buffer"/>. Единовременно допустим запуск только одной операции чтения.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="buffer">Данные, полученные от устройства.</param>
            ''' <param name="bytesToRead">Число байтов, которые нужно прочитать.</param>
            ''' <param name="bytesWereRed">Число прочитанных байтов.</param>
            ''' <remarks>Чтение заканчивается, если: получен индикатор конца передачи; прочитан символ окончания передачи; прочитанное число данных равно <paramref name="bytesToRead"/>.</remarks>
            <DllImport(LibPath, EntryPoint:="#256", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viRead(vi As Integer, buffer As StringBuilder, bytesToRead As Integer, <Out()> ByRef bytesWereRed As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Асинхронно читает данные с устройства.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="buffer">Данные, полученные от устройства.</param>
            ''' <param name="bytesToRead">Число байтов, которые нужно прочитать.</param>
            ''' <param name="jobId">Идентификатор задания данной асинхронной операции.</param>
            <DllImport(LibPath, EntryPoint:="#277", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viReadAsync(vi As Integer, buffer As Byte(), bytesToRead As Integer, <Out()> ByRef jobId As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Асинхронно читает данные с устройства.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="buffer">Данные, полученные от устройства.</param>
            ''' <param name="bytesToRead">Число байтов, которые нужно прочитать.</param>
            ''' <param name="jobId">Идентификатор задания данной асинхронной операции.</param>
            <DllImport(LibPath, EntryPoint:="#277", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viReadAsync(vi As Integer, buffer As StringBuilder, bytesToRead As Integer, <Out()> ByRef jobId As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Синхронно читает данные и сохраняет в файл.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="filename">Имя файла.</param>
            ''' <param name="bytesToRead">Число байтов, которые нужно прочитать.</param>
            ''' <param name="bytesWereRed">Число прочитанных байтов.</param>
            ''' <remarks>Если установлен атрибут <see cref="ViAttr.VI_ATTR_FILE_APPEND_EN"/>, данные будут добавляться к файлу, иначе - перезаписывать его.</remarks>
            <DllImport(LibPath, EntryPoint:="#219", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viReadToFile(vi As Integer, filename As String, bytesToRead As Integer, <Out()> ByRef bytesWereRed As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Синхронно записывает данные в устройство. Передаваемые данные это <paramref name="buffer"/>. 
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="buffer">Буфер для передачи.</param>
            ''' <param name="bytesToWrite">Сколько байтов нужно передать.</param>
            ''' <param name="bytesWereWritten">Сколько байтов было передано. Если передать <see cref="VI_BOOL.VI_NULL"/>, сюда не будет возвращено значение.</param>
            <DllImport(LibPath, EntryPoint:="#257", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viWrite(vi As Integer, buffer As Byte(), bytesToWrite As Integer, <Out()> ByRef bytesWereWritten As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Синхронно записывает данные в устройство. Передаваемые данные это <paramref name="buffer"/>. 
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="buffer">Буфер для передачи.</param>
            ''' <param name="bytesToWrite">Сколько байтов нужно передать. </param>
            ''' <param name="bytesWereWritten">Сколько байтов было передано. Если передать <see cref="VI_BOOL.VI_NULL"/>, сюда не будет возвращено значение.</param>
            <DllImport(LibPath, EntryPoint:="#257", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viWrite(vi As Integer, buffer As String, bytesToWrite As Integer, <Out()> ByRef bytesWereWritten As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Асинхронно записывает данные в устройство. Передаваемые данные это <paramref name="buffer"/>. 
            ''' </summary>
            ''' <param name="vi"></param>
            ''' <param name="buffer">Буфер для передачи.</param>
            ''' <param name="bytesToWrite">Сколько байтов нужно передать.</param>
            ''' <param name="jobId">Идентификатор задания данной асинхронной операции.</param>
            <DllImport(LibPath, EntryPoint:="#278", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viWriteAsync(vi As Integer, buffer As Byte(), bytesToWrite As Integer, <Out()> ByRef jobId As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Асинхронно записывает данные в устройство. Передаваемые данные это <paramref name="buffer"/>. 
            ''' </summary>
            ''' <param name="vi"></param>
            ''' <param name="buffer">Буфер для передачи.</param>
            ''' <param name="bytesToWrite">Сколько байтов нужно передать.</param>
            ''' <param name="jobId">Идентификатор задания данной асинхронной операции.</param>
            <DllImport(LibPath, EntryPoint:="#278", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viWriteAsync(vi As Integer, buffer As String, bytesToWrite As Integer, <Out()> ByRef jobId As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Берёт данные из файла и синхронно передаёт в устройство. 
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="filename">Имя файла.</param>
            ''' <param name="bytesToWrite">Сколько байтов нужно передать.</param>
            ''' <param name="bytesWereWritten">Сколько байтов было передано. Если передать <see cref="VI_BOOL.VI_NULL"/>, сюда не будет возвращено значение.</param>
            <DllImport(LibPath, EntryPoint:="#218", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viWriteFromFile(vi As Integer, filename As String, bytesToWrite As Integer, <Out()> ByRef bytesWereWritten As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Устанавливает программное или аппаратное прерывание. 
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="protocol">Протокол в процессе прерывания. Зависит от <see cref="ViAttr.VI_ATTR_TRIG_ID"/>, а также от <see cref="ViAttr.VI_ATTR_IO_PROT"/>.</param>
            <DllImport(LibPath, EntryPoint:="#258", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viAssertTrigger(vi As Integer, protocol As TrigProt) As ViStatus
            End Function

            ''' <summary>
            ''' Читает байт статуса запроса на обслуживание (команда *STB?\n).
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="status">Байт статуса запроса на обслуживание.</param>
            <DllImport(LibPath, EntryPoint:="#259", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viReadSTB(vi As Integer, <Out()> ByRef status As Short) As ViStatus
            End Function

            ''' <summary>
            ''' Очищает устройство (команда *CLS\n). 
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            <DllImport(LibPath, EntryPoint:="#260", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viClear(vi As Integer) As ViStatus
            End Function

#End Region '/БАЗОВЫЕ ОПЕРАЦИИ ВВОДА-ВЫВОДА

#Region "ОПЕРАЦИИ С ПАМЯТЬЮ"

            ''' <summary>
            ''' Выделяет память, которая была выделена в сессии, и возвращает смещение в памяти устройства <paramref name="offset"/>.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="memSize">Выделяемый размер памяти.</param>
            ''' <param name="offset">Указатель на выделенный участок памяти.</param>
            <DllImport(LibPath, EntryPoint:="#291", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viMemAlloc(vi As Integer, memSize As Integer, <Out()> ByRef offset As UIntPtr) As ViStatus
            End Function

            ''' <summary>
            ''' Очищает память, ранее выделенную <see cref="viMemAlloc(Integer, Integer, ByRef UIntPtr)"/>.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="offset">Задаёт ранее выделенную память.</param>
            <DllImport(LibPath, EntryPoint:="#292", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viMemFree(vi As Integer, offset As UIntPtr) As ViStatus
            End Function

#End Region '/ОПЕРАЦИИ С ПАМЯТЬЮ

#Region "СПЕЦИФИЧНЫЕ ДЛЯ ИНТЕРФЕЙСА ОПЕРАЦИИ"

            ''' <summary>
            ''' Управляет состоянием линии REN интерфейса GPIB. 
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="mode">Состояние линии REN и, опционально, локальное/удалённое состояние устройства.</param>
            <DllImport(LibPath, EntryPoint:="#208", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGpibControlREN(vi As Integer, mode As GpibRen) As ViStatus
            End Function

            ''' <summary>
            ''' Управляет состоянием линии ATN интерфейса GPIB и, опционально, состоянием активного контроллера локальной интерфейсной платы. 
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="mode">Состояние линии ATN интерфейса GPIB и, опционально, состоянием активного контроллера локальной интерфейсной платы. </param>
            <DllImport(LibPath, EntryPoint:="#210", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGpibControlATN(vi As Integer, mode As GpibAtn) As ViStatus
            End Function

            ''' <summary>
            ''' "Дёргает" линию очистки интерфейса (interface clear line, IFC) на промежуток как минимум 100 мкс.
            ''' Эта операция выставляет линию IFC и переходит в состояние controller in charge (CIC).
            ''' Локальная плата должна быть контроллером системы.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <remarks>Допустимо только в сессиях GPIB.</remarks>
            <DllImport(LibPath, EntryPoint:="#211", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGpibSendIFC(vi As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Записывает команду GPIB на шину.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="buffer">Буфер для передачи.</param>
            ''' <param name="bytesToWrite">Сколько байтов нужно передать.</param>
            ''' <param name="bytesWereWritten">Сколько байтов было передано. Если передать <see cref="VI_BOOL.VI_NULL"/>, сюда не будет возвращено значение.</param>
            <DllImport(LibPath, EntryPoint:="#212", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGpibCommand(vi As Integer, buffer As String, bytesToWrite As Integer, <Out()> bytesWereWritten As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Указание устройству GPIB по заданному адресу стать контроллером (controller in charge, CIC).
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="primAddr">Первичный адрес устройства, которому необходимо передать управление.</param>
            ''' <param name="secAddr">Вторичный адрес GPIB устройства. Если вторичного адреса нет, следует передать <see cref="GpibSecAddress.VI_NO_SEC_ADDR"/>.</param>
            <DllImport(LibPath, EntryPoint:="#213", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viGpibPassControl(vi As Integer, primAddr As Short, secAddr As GpibSecAddress) As ViStatus
            End Function

            ''' <summary>
            ''' Отправляет устройству команду или запрос, и/или получение ответа на предыдущий запрос.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="mode">Режим команды или запроса.</param>
            ''' <param name="devCmd">Команда для отправки.</param>
            ''' <param name="devResponse">Ответ устройства.</param>
            <DllImport(LibPath, EntryPoint:="#209", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viVxiCommandQuery(vi As Integer, mode As VxiCmd, devCmd As Integer, <Out()> ByRef devResponse As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Выставить специальный обслуживающий сигнал.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="line">Тип специального запроса.</param>
            <DllImport(LibPath, EntryPoint:="#214", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viAssertUtilSignal(vi As Integer, line As UtilAssert) As ViStatus
            End Function

            ''' <summary>
            ''' Выставляет специальное прерывание или сигнал.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="mode">Тип сигнала.</param>
            ''' <param name="statusId">Статус, который будет получен при прерывании цикла опроса.</param>
            <DllImport(LibPath, EntryPoint:="#215", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viAssertIntrSignal(vi As Integer, mode As ViAssert, statusId As Integer) As ViStatus
            End Function

            ''' <summary>
            ''' Привязывает заданный источник прерывания. Используется для связывания одной линии прерывания с другой.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="trigSrc">Линия-источник прерывания.</param>
            ''' <param name="trigDest">Линия-получатель прерывания.</param>
            ''' <param name="mode">Режим привязки. </param>
            <DllImport(LibPath, EntryPoint:="#216", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viMapTrigger(vi As Integer, trigSrc As TrigId, trigDest As TrigId, mode As TrigId) As ViStatus
            End Function

            ''' <summary>
            ''' Используется для привязки одной линии прерывания к другой.
            ''' </summary>
            ''' <param name="vi">Логический идентификатор сессии.</param>
            ''' <param name="trigSrc">Линия-источник прерывания с предыдущей привязки. </param>
            ''' <param name="trigDest">Линия-получатель прерывания с предыдущей привязки.</param>
            <DllImport(LibPath, EntryPoint:="#217", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viUnmapTrigger(vi As Integer, trigSrc As TrigId, trigDest As TrigId) As ViStatus
            End Function

            <DllImport(LibPath, EntryPoint:="#293", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viUsbControlOut(vi As Integer, bmRequestType As Short, bRequest As Short, wValue As Short, wIndex As Short, wLength As Short, buf As Byte()) As ViStatus
            End Function

            <DllImport(LibPath, EntryPoint:="#294", ExactSpelling:=True, CharSet:=CharSet.Ansi)>
            Public Shared Function viUsbControlIn(vi As Integer, mRequestType As Short, bRequest As Short, wValue As Short, wIndex As Short, wLength As Short, buf As Byte(), <Out()> ByRef retCnt As Short) As ViStatus
            End Function

#End Region '/СПЕЦИФИЧНЫЕ ДЛЯ ИНТЕРФЕЙСА ОПЕРАЦИИ

#Region "CONST"

            ''' <summary>
            ''' Недействительный идентификатор сессии.
            ''' </summary>
            Public Const VI_ERROR_INV_SESSION As Integer = -1073807346
            Public Const VI_SPEC_VERSION As Integer = &H400000
            Public Const VI_FIND_BUFLEN As Short = &H100
            Public Const VI_UNKNOWN_LA As Short = -1S

#End Region '/CONST

#Region "ENUM"

            Public Enum Prot As Short
                VI_PROT_NORMAL = 1S
                VI_PROT_FDC = 2S
                VI_PROT_HS488 = 3S
                VI_PROT_4882_STRS = 4S
                VI_PROT_USBTMC_VENDOR = 5S
            End Enum

            ''' <summary>
            ''' Режим бфстрого канала данных (fast data channel, FDC).
            ''' </summary>
            Public Enum FdcMode As Short
                ''' <summary>
                ''' Обычный режим.
                ''' </summary>
                VI_FDC_NORMAL = 1S
                ''' <summary>
                ''' Потоковый режим.
                ''' </summary>
                VI_FDC_STREAM = 2S
            End Enum

            Public Enum MemorySpace As Short
                ''' <summary>
                ''' Неизвестно.
                ''' </summary>
                VI_OPAQUE_SPACE = -1S
                ''' <summary>
                ''' Локальная память процесса (используется виртуальный адрес).
                ''' </summary>
                VI_LOCAL_SPACE = 0S
                ''' <summary>
                ''' Адресное пространство A16 шины VXI/MXI.
                ''' </summary>
                VI_A16_SPACE = 1S
                ''' <summary>
                ''' Адресное пространство A24 шины VXI/MXI.
                ''' </summary>
                VI_A24_SPACE = 2S
                ''' <summary>
                ''' Адресное пространство A32 шины VXI/MXI.
                ''' </summary>
                VI_A32_SPACE = 3S
            End Enum

            Public Enum Slot As Short
                ''' <summary>
                ''' Физический номер слота устройства неизвестен.
                ''' </summary>
                VI_UNKNOWN_SLOT = -1S
                SLOT_0 = 0
                SLOT_1
                SLOT_2
                SLOT_3
                SLOT_4
                SLOT_5
                SLOT_6
                SLOT_7
                SLOT_8
                SLOT_9
                SLOT_10
                SLOT_11
                SLOT_12
            End Enum

            ''' <summary>
            ''' Уровни срабатывания прерываний.
            ''' </summary>
            Public Enum Levels As Short
                VI_UNKNOWN_LEVEL = -1S
                LEVEL_1 = 1
                LEVEL_2
                LEVEL_3
                LEVEL_4
                LEVEL_5
                LEVEL_6
                LEVEL_7
            End Enum

            <Flags()>
            Public Enum EventMechanism As Short
                VI_ALL_MECH = -1S
                ''' <summary>
                ''' Сессия ожидает событие. 
                ''' События должны запрашиваться вручную через функцию <see cref="viWaitOnEvent(Integer, ViEventType, Integer, ByRef ViEventType, ByRef Integer)"/>.
                ''' </summary>
                VI_QUEUE = 1S
                ''' <summary>
                ''' Сессия получает события с помощью функций обратного вызова по указателю, 
                ''' которые устанавливаются с помощью <see cref="viInstallHandler(Integer, ViEventType, ViEventHandler, Integer)"/>.
                ''' </summary>
                VI_HNDLR = 2S
                ''' <summary>
                ''' Сессия получает события с помощью функций обратного вызова по запросу.
                ''' События не будут доставлены сессии, пока <see cref="viEnableEvent(Integer, ViEventType, EventMechanism, Integer)"/> не будет вызвано снова через механизм VI_HNDLR.
                ''' </summary>
                VI_SUSPEND_HNDLR = 4S
            End Enum

            ''' <summary>
            ''' Протокол прерывания.
            ''' </summary>
            Public Enum TrigProt As Short
                ''' <summary>
                ''' То же, что и <see cref="VI_TRIG_PROT_SYNC"/>.
                ''' </summary>
                VI_TRIG_PROT_DEFAULT = 0S
                ''' <summary>
                ''' Устанавливает прерывание.
                ''' </summary>
                VI_TRIG_PROT_ON = 1S
                ''' <summary>
                ''' Отключает прерывание.
                ''' </summary>
                VI_TRIG_PROT_OFF = 2S
                ''' <summary>
                ''' Импульс (последовательно установка и снятие прерывания).
                ''' </summary>
                VI_TRIG_PROT_SYNC = 5S
            End Enum

            Public Enum ViIo As Short
                ''' <summary>
                ''' Форматированный буфер чтения.
                ''' </summary>
                VI_READ_BUF = 1S
                ''' <summary>
                ''' Форматированный буфер записи.
                ''' </summary>
                VI_WRITE_BUF = 2S
                ''' <summary>
                ''' Сбрасывает содержимое буфера чтения (не позволяет совершать операции ввода-вывода с устройством).
                ''' </summary>
                VI_READ_BUF_DISCARD = 4S
                ''' <summary>
                ''' Сбрасывает содержимое буфера записи (не позволяет совершать операции ввода-вывода с устройством).
                ''' </summary>
                VI_WRITE_BUF_DISCARD = 8S
                ''' <summary>
                ''' Сбрасывает содержимое буфера приёма (то же, что <see cref="VI_IO_IN_BUF_DISCARD"/>).
                ''' </summary>
                VI_IO_IN_BUF = 16S
                ''' <summary>
                ''' Очищает буфер передачи, записывая все данные из буфера в устройство.
                ''' </summary>
                VI_IO_OUT_BUF = 32S
                ''' <summary>
                ''' Сбрасывает содержимое буфера приёма (не позволяет совершать операции ввода-вывода с устройством).
                ''' </summary>
                VI_IO_IN_BUF_DISCARD = 64S
                ''' <summary>
                ''' Сбрасывает содержимое буфера передачи (не позволяет совершать операции ввода-вывода с устройством).
                ''' </summary>
                VI_IO_OUT_BUF_DISCARD = 128S
            End Enum

            Public Enum WinAccess As Short
                ''' <summary>
                ''' Через привязку.
                ''' </summary>
                VI_NMAPPED = 1S
                ''' <summary>
                ''' Через операции.
                ''' </summary>
                VI_USE_OPERS = 2S
                ''' <summary>
                ''' Напрямую, отменяя ссылку на адреса.
                ''' </summary>
                VI_DEREF_ADDR = 3S
            End Enum

            Public Enum Timeout As Integer
                ''' <summary>
                ''' Отсутствие таймаута.
                ''' </summary>
                VI_TMO_IMMEDIATE = 0
                ''' <summary>
                ''' Отключает механизм таймаута.
                ''' </summary>
                VI_TMO_INFINITE = -1
            End Enum

            Public Enum Parity As Short
                ''' <summary>
                ''' Проверка чётности отсутствует.
                ''' </summary>
                VI_ASRL_PAR_NONE = 0S
                ''' <summary>
                ''' Проверка нечётности.
                ''' </summary>
                VI_ASRL_PAR_ODD = 1S
                ''' <summary>
                ''' Проверка чётности.
                ''' </summary>
                VI_ASRL_PAR_EVEN = 2S
                ''' <summary>
                ''' Бит чётности всегда "1".
                ''' </summary>
                VI_ASRL_PAR_MARK = 3S
                ''' <summary>
                ''' Бит чётности всегда "0".
                ''' </summary>
                VI_ASRL_PAR_SPACE = 4S
            End Enum

            Public Enum StopBits As Short
                ''' <summary>
                ''' Один стоповый бит.
                ''' </summary>
                VI_ASRL_STOP_ONE = 10S
                ''' <summary>
                ''' Полтора стоповых бита.
                ''' </summary>
                VI_ASRL_STOP_ONE5 = 15S
                ''' <summary>
                ''' Два стоповых бита.
                ''' </summary>
                VI_ASRL_STOP_TWO = 20S
            End Enum

            ''' <summary>
            ''' Контроль потока.
            ''' </summary>
            ''' <remarks>
            ''' Можно использовать несколько механизмов одновременно, если их поддерживает аппаратура.
            ''' </remarks>
            <Flags>
            Public Enum FlowCntrl As Short
                ''' <summary>
                ''' Буферы на обоих концах соединения должны быть достаточного размера, чтобы вместить все данные. 
                ''' </summary>
                VI_ASRL_FLOW_NONE = 0S
                ''' <summary>
                ''' Для контроля потока используются символы XON и XOFF. 
                ''' Когда буфер приёма почти полон, он выставляет XOFF, и передающая сторона приостанавливает передачу.
                ''' </summary>
                VI_ASRL_FLOW_XON_XOFF = 1S
                ''' <summary>
                ''' Для контроля потока используются сигналы RTS и CTS. 
                ''' Передающий механизм контролирует входящий поток снятием RTS, когда приёмный буфер почти полон. 
                ''' Он контролирует исходящий поток приостановкой передачи, когда снимается CTS.
                ''' </summary>
                VI_ASRL_FLOW_RTS_CTS = 2S
                ''' <summary>
                ''' Для контроля потока используются сигналы DTR и DSR. 
                ''' Передающий механизм контролирует входящий поток снятием сигнала DTR, когда приёмный буфер почти полон. 
                ''' Он контролирует исходящий поток приостановкой передачи, когда снимается DSR.
                ''' </summary>
                VI_ASRL_FLOW_DTR_DSR = 4S
            End Enum

            Public Enum Ending As Short
                ''' <summary>
                ''' Передача не будет завершена, пока все данные не будут переданы или не возникнет ошибка.
                ''' </summary>
                VI_ASRL_END_NONE = 0S
                ''' <summary>
                ''' Будет переданы все данные, кроме последнего символа со сброшенным последним битом. 
                ''' Затем будет передан последний символ с установленным последним битом.
                ''' </summary>
                VI_ASRL_END_LAST_BIT = 1S
                ''' <summary>
                ''' Передача данных будет завершаться терминальным символом.
                ''' </summary>
                VI_ASRL_END_TERMCHAR = 2S
                ''' <summary>
                ''' Передача данных будет завершаться символом прерывания.
                ''' </summary>
                VI_ASRL_END_BREAK = 3S
            End Enum

            Public Enum Endianess As Short
                VI_BIG_ENDIAN = 0S
                VI_LITTLE_ENDIAN = 1S
            End Enum

            Public Enum PrivateAccess As Short
                VI_DATA_PRIV = 0S
                VI_DATA_NPRIV = 1S
                VI_PROG_PRIV = 2S
                VI_PROG_NPRIV = 3S
                VI_BLCK_PRIV = 4S
                VI_BLCK_NPRIVt = 5S
                VI_D64_PRIV = 6S
                VI_D64_NPRIV = 7S
            End Enum

            Public Enum ViWidth As Short
                ''' <summary>
                ''' 8-битная передача.
                ''' </summary>
                VI_WIDTH_8 = 1S
                ''' <summary>
                ''' 16-битная передача.
                ''' </summary>
                VI_WIDTH_16 = 2S
                ''' <summary>
                ''' 32-битная передача.
                ''' </summary>
                VI_WIDTH_32 = 4S
            End Enum

            Public Enum GpibRen As UShort
                ''' <summary>
                ''' Снять линию REN.
                ''' </summary>
                VI_GPIB_REN_DEASSERT = 0S
                ''' <summary>
                ''' Выставить линию REN.
                ''' </summary>
                VI_GPIB_REN_ASSERT = 1S
                ''' <summary>
                ''' Послать команду Go to local (GTL) этому устройству и снять линию REN.
                ''' </summary>
                VI_GPIB_REN_DEASSERT_GTL = 2S
                ''' <summary>
                ''' Установить линию REN и адрес этого устройства.
                ''' </summary>
                VI_GPIB_REN_ASSERT_ADDRESS = 3S
                ''' <summary>
                ''' Послать LLO любым слушающим устройствам.
                ''' </summary>
                VI_GPIB_REN_ASSERT_LLO = 4S
                ''' <summary>
                ''' Обратиться к этому устройству и послать LLO, установить его в RWLS.
                ''' </summary>
                VI_GPIB_REN_ASSERT_ADDRESS_LLO = 5S
                ''' <summary>
                ''' Послать команду Go to local (GTL) этому устройству.
                ''' </summary>
                VI_GPIB_REN_ADDRESS_GTL = 6S
            End Enum

            Public Enum GpibAtn As UShort
                ''' <summary>
                ''' Снять линию ATN.
                ''' </summary>
                VI_GPIB_ATN_DEASSERT = 0S
                ''' <summary>
                ''' Выставляет синхронно линию ATN. Если происходит рукопожатие, ATN не будет установлено до его окончания.
                ''' </summary>
                VI_GPIB_ATN_ASSERT = 1S
                ''' <summary>
                ''' Снять линию ATN и войти в режим скрытного рукопожатия. 
                ''' Локальная плата переходит в режим рукопожатия как принимающая сторона без действительного чтения данных.
                ''' </summary>
                VI_GPIB_ATN_DEASSERT_HANDSHAKE = 2S
                ''' <summary>
                ''' Установить линию ATN асинхронно. Обычно используется во время возникновения ошибок.
                ''' </summary>
                VI_GPIB_ATN_ASSERT_IMMEDIATE = 3S
            End Enum

            Public Enum GpibSpeed As Short
                VI_GPIB_HS488_NIMPL = -1S
                VI_GPIB_HS488_DISABLED = 0S
            End Enum

            Public Enum GpibRole As Short
                VI_GPIB_UNADDRESSED = 0S
                VI_GPIB_TALKER = 1S
                VI_GPIB_LISTENER = 2S
            End Enum

            Public Enum VxiCmd As Short
                ''' <summary>
                ''' Отправить 16-битную команду.
                ''' </summary>
                VI_VXI_CMD16 = &H200
                ''' <summary>
                ''' Отправить 16-битную команду, получить 16-битный ответ.
                ''' </summary>
                VI_VXI_CMD16_RESP16 = &H202
                ''' <summary>
                ''' Получить 16-битный ответ на предыдущий запрос.
                ''' </summary>
                VI_VXI_RESP16 = &H2
                ''' <summary>
                ''' Отправить 32-битную команду.
                ''' </summary>
                VI_VXI_CMD32 = &H400
                ''' <summary>
                ''' Отправить 32-битную команду, получить 16-битный ответ.
                ''' </summary>
                VI_VXI_CMD32_RESP16 = &H402
                ''' <summary>
                ''' Отправить 32-битную команду, получить 32-битный ответ.
                ''' </summary>
                VI_VXI_CMD32_RESP32 = &H404
                ''' <summary>
                ''' Получить 32-битный ответ на предыдущий запрос.
                ''' </summary>
                VI_VXI_RESP32 = &H4
            End Enum

            Public Enum ViAssert As Short
                ''' <summary>
                ''' Послать уведомлениее через сигнал VXI.
                ''' </summary>
                VI_ASSERT_SIGNAL = -1S
                ''' <summary>
                ''' Использовать любое уведомление, назначенное данному устройству.
                ''' </summary>
                VI_ASSERT_USE_ASSIGNED = 0S
                ''' <summary>
                ''' Послать прерывание через назначенную линию IRQ1.
                ''' </summary>
                VI_ASSERT_IRQ1 = 1S
                ''' <summary>
                ''' Послать прерывание через назначенную линию IRQ2.
                ''' </summary>
                VI_ASSERT_IRQ2 = 2S
                ''' <summary>
                ''' Послать прерывание через назначенную линию IRQ3.
                ''' </summary>
                VI_ASSERT_IRQ3 = 3S
                ''' <summary>
                ''' Послать прерывание через назначенную линию IRQ4.
                ''' </summary>
                VI_ASSERT_IRQ4 = 4S
                ''' <summary>
                ''' Послать прерывание через назначенную линию IRQ5.
                ''' </summary>
                VI_ASSERT_IRQ5 = 5S
                ''' <summary>
                ''' Послать прерывание через назначенную линию IRQ6.
                ''' </summary>
                VI_ASSERT_IRQ6 = 6S
                ''' <summary>
                ''' Послать прерывание через назначенную линию IRQ7.
                ''' </summary>
                VI_ASSERT_IRQ7 = 7S
            End Enum

            Public Enum UtilAssert As Short
                ''' <summary>
                ''' Выставить сброс системы.
                ''' </summary>
                VI_UTIL_ASSERT_SYSRESET = 1S
                ''' <summary>
                ''' Выставить сбой системы.
                ''' </summary>
                VI_UTIL_ASSERT_SYSFAIL = 2S
                ''' <summary>
                ''' Убрать сбой системы.
                ''' </summary>
                VI_UTIL_DEASSERT_SYSFAIL = 3S
            End Enum

            Public Enum DeviceClass As Short
                ''' <summary>
                ''' Память.
                ''' </summary>
                VI_VXI_CLASS_MEMORY = 0S
                ''' <summary>
                ''' Расширенный.
                ''' </summary>
                VI_VXI_CLASS_EXTENDED = 1S
                ''' <summary>
                ''' Основан на сообщениях.
                ''' </summary>
                VI_VXI_CLASS_MESSAGE = 2S
                ''' <summary>
                ''' Основан на регистрах.
                ''' </summary>
                VI_VXI_CLASS_REGISTER = 3S
                ''' <summary>
                ''' Основан на регистрах или другой.
                ''' </summary>
                VI_VXI_CLASS_OTHER = 4S
            End Enum

            ''' <summary>
            ''' Добавлено для совместимости. Частично дублирует <see cref="ViIo"/>.
            ''' </summary>
            Public Enum ViAsrl As Short
                VI_ASRL488 = 4S
                VI_ASRL_IN_BUF = 16S
                VI_ASRL_OUT_BUF = 32S
                VI_ASRL_IN_BUF_DISCARD = 64S
                VI_ASRL_OUT_BUF_DISCARD = 128S
            End Enum

            Public Enum BufOperMode As Short
                ''' <summary>
                ''' Буфер очищается каждый раз, когда операция завершается.
                ''' </summary>
                VI_FLUSH_ON_ACCESS = 1S
                ''' <summary>
                ''' Буфер очищается, когда полон, или когда получен индикатор END.
                ''' </summary>
                VI_FLUSH_WHEN_FULL = 2S
                ''' <summary>
                ''' Буфер очищается только по запросу.
                ''' </summary>
                VI_FLUSH_DISABLE = 3S
            End Enum

            Public Enum VI_BOOL As UShort
                VI_NULL = 0US
                VI_FALSE = 0US
                VI_TRUE = 1US
            End Enum

            Public Enum IoProt As Short
                VI_NORMAL = 1S
                VI_FDC = 2S
                VI_HS488 = 3S
            End Enum

            ''' <summary>
            ''' Режим доступа к ресурсу.
            ''' </summary>
            <Flags()>
            Public Enum AccessMode As Short
                ''' <summary>
                ''' Без блокировки.
                ''' </summary>
                VI_NO_LOCK = 0S
                ''' <summary>
                ''' Получает эксклюзивную блокировку ресурса в данной сессии.
                ''' </summary>
                VI_EXCLUSIVE_LOCK = 1S
                ''' <summary>
                ''' 
                ''' </summary>
                VI_SHARED_LOCK = 2S
                ''' <summary>
                ''' Конфигурирование атрибутов, заданных внешнией утилитой конфигурации.
                ''' </summary>
                VI_LOAD_CONFIG = 4S
            End Enum

            ''' <summary>
            ''' Режим прерывания.
            ''' </summary>
            Public Enum TrigId As Short
                VI_TRIG_ALL = -2S
                VI_TRIG_SW = -1S
                VI_TRIG_TTL0 = 0S
                VI_TRIG_TTL1 = 1S
                VI_TRIG_TTL2 = 2S
                VI_TRIG_TTL3 = 3S
                VI_TRIG_TTL4 = 4S
                VI_TRIG_TTL5 = 5S
                VI_TRIG_TTL6 = 6S
                VI_TRIG_TTL7 = 7S
                VI_TRIG_ECL0 = 8S
                VI_TRIG_ECL1 = 9S
                VI_TRIG_PANEL_IN = 27S
                VI_TRIG_PANEL_OUT = 28S
            End Enum

            ''' <summary>
            ''' Вторичный адрес GPIB контроллера.
            ''' </summary>
            Public Enum GpibSecAddress As Short
                ''' <summary>
                ''' У устройства нет вторичного адреса.
                ''' </summary>
                VI_NO_SEC_ADDR = -1S
                ADDR_0 = 0
                ADDR_1
                ADDR_2
                ADDR_3
                ADDR_4
                ADDR_5
                ADDR_6
                ADDR_7
                ADDR_8
                ADDR_9
                ADDR_10
                ADDR_11
                ADDR_12
                ADDR_13
                ADDR_14
                ADDR_15
                ADDR_16
                ADDR_17
                ADDR_18
                ADDR_19
                ADDR_20
                ADDR_21
                ADDR_22
                ADDR_23
                ADDR_24
                ADDR_25
                ADDR_26
                ADDR_27
                ADDR_28
                ADDR_29
                ADDR_30
                ADDR_31
            End Enum

            Public Enum State As Short
                VI_STATE_UNKNOWN = -1S
                VI_STATE_UNASSERTED = 0S
                VI_STATE_ASSERTED = 1S
            End Enum

#End Region '/ENUMS

        End Class '/Native

#Region "ENUMS"

        ''' <summary>
        ''' Attributes.
        ''' </summary>
        Public Enum ViAttr As Integer
            ''' <summary>
            ''' Класс ресурса.
            ''' </summary>
            VI_ATTR_RSRC_CLASS = -1073807359
            ''' <summary>
            ''' Имя ресурса.
            ''' </summary>
            VI_ATTR_RSRC_NAME
            ''' <summary>
            ''' Версия реализации ресурса.
            ''' </summary>
            VI_ATTR_RSRC_IMPL_VERSION = 1073676291
            ''' <summary>
            ''' Состояние блокировки ресурса в текущей сессии.
            ''' </summary>
            VI_ATTR_RSRC_LOCK_STATE
            ''' <summary>
            ''' Максимальная длина очереди.
            ''' </summary>
            VI_ATTR_MAX_QUEUE_LENGTH
            ''' <summary>
            ''' Пользовательские данные.
            ''' </summary>
            VI_ATTR_USER_DATA = 1073676295
            ''' <summary>
            ''' Какой канал FDC (0..7) будет использован для передачи.
            ''' </summary>
            VI_ATTR_FDC_CHNL = 1073676301
            ''' <summary>
            ''' Используемый режим FDC (fast data channel). Допустимые значения <see cref="Native.FdcMode"/>.
            ''' </summary>
            VI_ATTR_FDC_MODE = 1073676303
            ''' <summary>
            ''' Значение <see cref="Native.VI_BOOL.VI_TRUE"/> позволяет обслуживающему посылать сигнал, когда управление каналом FDC возвращается контроллеру.
            ''' Это действие освобождает контроллер от необходимости опроса заголовка FDC в процессе передачи данных.
            ''' </summary>
            VI_ATTR_FDC_GEN_SIGNAL_EN = 1073676305
            ''' <summary>
            ''' Если <see cref="Native.VI_BOOL.VI_TRUE"/>, для передачи данных будет использована пара каналов. Иначе будет использован только один канал.
            ''' </summary>
            VI_ATTR_FDC_USE_PAIR = 1073676307
            ''' <summary>
            ''' Вставлять ли END после время передачи последнего байта в буфере.
            ''' </summary>
            VI_ATTR_SEND_END_EN = 1073676310
            ''' <summary>
            ''' Символ конца передачи.
            ''' </summary>
            VI_ATTR_TERMCHAR = 1073676312
            ''' <summary>
            ''' Минимальное значение таймаута, мс. Может принимать значение из <see cref="native.Timeout"/>.
            ''' </summary>
            VI_ATTR_TMO_VALUE = 1073676314
            ''' <summary>
            ''' Использовать ли повтор адресации перед каждой операцией передачи данных. Значения <see cref="Native.VI_BOOL"/>.
            ''' </summary>
            VI_ATTR_GPIB_READDR_EN
            ''' <summary>
            ''' Задаёт используемый протокол. Значения могут быть <see cref="Native.IoProt"/>.
            ''' </summary>
            ''' <remarks>
            ''' Например, в системах VXI можно выбрать м/у нормальным словным режимом и быстрым (fast data channel, FDC). 
            ''' В GPIB можно выбрать м/у нормальным и высокоскоростным (high-speed, HS488) режимами. 
            ''' В системах ASRL можно выбрать м/у нормальным и передачей типа 488, в случае чего операции viAssertTrigger/viReadSTB/viClear посылают 488.2-совместимые строки.
            ''' </remarks>
            VI_ATTR_IO_PROT
            ''' <summary>
            ''' Задаёт, при операциях ввода-вывода будет использован DMA или программируемый I/O. Значения <see cref="Native.VI_BOOL"/>.
            ''' </summary>
            VI_ATTR_DMA_ALLOW_EN = 1073676318
            ''' <summary>
            ''' Скорость интерфейса.
            ''' </summary>
            ''' <remarks>
            ''' Можно задать 32-разрядное целое число, но обычно используются стандартные значения 300, 1200, 2400 или 9600.
            ''' </remarks>
            VI_ATTR_ASRL_BAUD = 1073676321
            ''' <summary>
            ''' Число битов данных в каждом кадре (5..8).
            ''' </summary>
            VI_ATTR_ASRL_DATA_BITS
            ''' <summary>
            ''' Проверка чётности при передаче каждого кадра. Может принимать значения из <see cref="Native.Parity"/>.
            ''' </summary>
            VI_ATTR_ASRL_PARITY
            ''' <summary>
            ''' Число стоповых бит, показывающих конец кадра. Может принимать значения из <see cref="Native.StopBits"/>.
            ''' </summary>
            VI_ATTR_ASRL_STOP_BITS
            ''' <summary>
            ''' Использовать ли контроль потока. Может принимать значения из <see cref="Native.FlowCntrl"/>. 
            ''' </summary>
            VI_ATTR_ASRL_FLOW_CNTRL
            ''' <summary>
            ''' Определяет режим работы буфера чтения. Значения могут быть <see cref="Native.BufOperMode"/>.
            ''' </summary>
            VI_ATTR_RD_BUF_OPER_MODE = 1073676330
            ''' <summary>
            ''' Размер буфера чтения.
            ''' </summary>
            VI_ATTR_RD_BUF_SIZE
            ''' <summary>
            ''' Режим буфера записи. Значения могут быть <see cref="Native.BufOperMode"/>.
            ''' </summary>
            VI_ATTR_WR_BUF_OPER_MODE = 1073676333
            ''' <summary>
            ''' Размер буфера записи.
            ''' </summary>
            VI_ATTR_WR_BUF_SIZE
            ''' <summary>
            ''' Подавлять ли индикатор END. Если задано <see cref="Native.VI_BOOL.VI_TRUE"/>, то END не будет завершать операции чтения.
            ''' </summary>
            VI_ATTR_SUPPRESS_END_EN = 1073676342
            ''' <summary>
            ''' Использовать ли символ конца передачи. Значения <see cref="Native.VI_BOOL.VI_TRUE"/>.
            ''' </summary>
            VI_ATTR_TERMCHAR_EN = 1073676344
            ''' <summary>
            ''' Задаёт модификатор адреса для применения в высокоуровневых операциях доступа. Может принимать значения из <see cref="Native.PrivateAccess"/>.
            ''' </summary>
            VI_ATTR_DEST_ACCESS_PRIV
            ''' <summary>
            ''' Задаёт порядок байтов при передаче получателю. Может принимать значения из <see cref="Native.Endianess"/>.
            ''' </summary>
            VI_ATTR_DEST_BYTE_ORDER
            ''' <summary>
            ''' Задаёт модификатор адреса при операциях чтения. Может принимать значения из <see cref="Native.PrivateAccess"/>.
            ''' </summary>
            VI_ATTR_SRC_ACCESS_PRIV = 1073676348
            ''' <summary>
            ''' Задаёт порядок байтов при чтении из источника. Может принимать значения из <see cref="Native.Endianess"/>.
            ''' </summary>
            VI_ATTR_SRC_BYTE_ORDER
            ''' <summary>
            ''' На сколько должно быть инкрементировано смещение источника после каждого чтения (0 или 1, по умолчанию 1).
            ''' </summary>
            VI_ATTR_SRC_INCREMENT = 1073676352
            ''' <summary>
            ''' На сколько должно быть инкрементировано смещение получателя после каждой записи (0 или 1, по умолчанию 1).
            ''' </summary>
            VI_ATTR_DEST_INCREMENT
            ''' <summary>
            ''' Задаёт модификатор адреса для низкоуровневых операций, когда получает доступ к назначенному окну. Может принимать значения из <see cref="Native.PrivateAccess"/>.
            ''' </summary>
            VI_ATTR_WIN_ACCESS_PRIV = 1073676357
            ''' <summary>
            ''' Задаёт порядок байтов в низкоуровневых операциях, когда получает доступ к назначенному окну. Может принимать значения из <see cref="Native.Endianess"/>.
            ''' </summary>
            VI_ATTR_WIN_BYTE_ORDER = 1073676359
            ''' <summary>
            ''' Состояние линии GPIB ATN (ATtentioN). Принимает значения <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_GPIB_ATN_STATE = 1073676375
            ''' <summary>
            ''' Показывает, заданный интерфейс GPIB адресован на передачу, приём или не адресован. Принимает значения <see cref="Native.GpibRole"/>.
            ''' </summary>
            VI_ATTR_GPIB_ADDR_STATE = 1073676380
            ''' <summary>
            ''' Показывает, заданный интерфейс GPIB в режиме CIC (controller in charge) или нет. Принимает значения <see cref="Native.VI_BOOL"/>.
            ''' </summary>
            VI_ATTR_GPIB_CIC_STATE = 1073676382
            ''' <summary>
            ''' Показывает состояние линии NDAC (Not Data ACcepted) заданного интерфейса GPIB. Принимает значения <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_GPIB_NDAC_STATE = 1073676386
            ''' <summary>
            ''' Показывает состояние линии SRQ (Service ReQuest) заданного интерфейса GPIB. Принимает значения <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_GPIB_SRQ_STATE = 1073676391
            ''' <summary>
            ''' Показывает, является ли заданный интерфейс GPIB контроллером. Принимает значения <see cref="Native.VI_BOOL"/>.
            ''' </summary>
            VI_ATTR_GPIB_SYS_CNTRL_STATE
            ''' <summary>
            ''' Задаёт длину кабеля GPIB в метрах. Принимает значения 1..15 или <see cref="Native.GpibSpeed"/>.
            ''' </summary>
            VI_ATTR_GPIB_HS488_CBL_LEN
            ''' <summary>
            ''' Логический адрес контроллера в текущей сессии. Может принимать значения 0..255 и <see cref="Native.VI_UNKNOWN_LA"/>.
            ''' </summary>
            VI_ATTR_CMDR_LA = 1073676395
            ''' <summary>
            ''' Класс устройства. Может принимать значения из <see cref="Native.DeviceClass"/>.
            ''' </summary>
            VI_ATTR_VXI_DEV_CLASS
            ''' <summary>
            ''' Логический адрес устройства в мейнфрейме. Обычно это устройство с наименьшим логическим адресом. Может принимать значения 0..255 и <see cref="Native.VI_UNKNOWN_LA"/>.
            ''' </summary>
            VI_ATTR_MAINFRAME_LA = 1073676400
            ''' <summary>
            ''' Название производителя.
            ''' </summary>
            VI_ATTR_MANF_NAME = -1073807246
            ''' <summary>
            ''' Название модели устройства.
            ''' </summary>
            VI_ATTR_MODEL_NAME = -1073807241
            ''' <summary>
            ''' Показывает состояние линий прерывания VXI/VME. Это вектор бит, в котором биты 0..6 показывают состояние линий прерывания 1..7.
            ''' </summary>
            VI_ATTR_VXI_VME_INTR_STATUS = 1073676427
            ''' <summary>
            ''' Показывает состояние линий прерывания VXI. 
            ''' Это битовый вектор, в котором биты 0..9 соответствуют от <see cref="native.TrigId.VI_TRIG_TTL0"/> до <see cref="native.TrigId.VI_TRIG_ECL1"/>.
            ''' </summary>
            VI_ATTR_VXI_TRIG_STATUS = 1073676429
            ''' <summary>
            ''' Показывает состояние линий VXI/VME SYSFAIL (SYStem FAILure). Принимает значения <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_VXI_VME_SYSFAIL_STATE = 1073676436
            ''' <summary>
            ''' Базовый адрес интерфейсной шины, к которой привязано окно.
            ''' </summary>
            VI_ATTR_WIN_BASE_ADDR = 1073676440
            ''' <summary>
            ''' Размер области, свзяанной с окном.
            ''' </summary>
            VI_ATTR_WIN_SIZE = 1073676442
            ''' <summary>
            ''' Показывает число байтов, доступных в глобальном приёмном буфере.
            ''' </summary>
            VI_ATTR_ASRL_AVAIL_NUM = 1073676460
            ''' <summary>
            ''' Базовый адрес устройства VXI в адресном пространстве.
            ''' </summary>
            VI_ATTR_MEM_BASE
            ''' <summary>
            ''' Показывает состояние сигнала CTS (Clear To Send). Может принимать значения из <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_ASRL_CTS_STATE
            ''' <summary>
            ''' Показывает состояние сигнала DCD (Data Carrier Detect, обнаружение источника (carrier) - удалённого модема - на линии). 
            ''' Может принимать значения из <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_ASRL_DCD_STATE
            ''' <summary>
            ''' Показывает состояние сигнала DSR (Data Set Ready). Может принимать значения из <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_ASRL_DSR_STATE = 1073676465
            ''' <summary>
            ''' Используется для ручной установки или снятия сигнала DTR (Data Terminal Ready). Может принимать значения из <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_ASRL_DTR_STATE
            ''' <summary>
            ''' Показывает способ прерывания операции чтения. Может принимать значения из <see cref="Native.Ending"/>.
            ''' </summary>
            VI_ATTR_ASRL_END_IN
            ''' <summary>
            ''' Показывает способ прерывания операции записи. Может принимать значения из <see cref="Native.Ending"/>.
            ''' </summary>
            VI_ATTR_ASRL_END_OUT
            ''' <summary>
            ''' Задаёт символ для замены входящих символов с ошибкой. Может принимать значения 0..FF.
            ''' </summary>
            VI_ATTR_ASRL_REPLACE_CHAR = 1073676478
            ''' <summary>
            ''' Показывает состояние сигнала RI (Ring Indicator). Сигнал RI часто используется модемами, чтобы показать звонок. Может принимать значения из <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_ASRL_RI_STATE
            ''' <summary>
            ''' Используется для ручной установки или снятия сигнала RTS (Ready To Send). Может принимать значения из <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_ASRL_RTS_STATE
            ''' <summary>
            ''' Задаёт значение символа XON при управлении потоком XON/XOFF. Может принимать значения 0..FF.
            ''' </summary>
            VI_ATTR_ASRL_XON_CHAR
            ''' <summary>
            ''' Задаёт значение символа XOFF при управлении потоком XON/XOFF. Может принимать значения 0..FF.
            ''' </summary>
            VI_ATTR_ASRL_XOFF_CHAR
            ''' <summary>
            ''' Режимы, в которые может быть установлено текущее окно. Может принимать значения из <see cref="Native.WinAccess"/>.
            ''' </summary>
            VI_ATTR_WIN_ACCESS
            ''' <summary>
            ''' Сессия менеджера ресурсов.
            ''' </summary>
            VI_ATTR_RM_SESSION
            ''' <summary>
            ''' Логический адрес устройства VXI или VME, используемого в сессии. Может принимать значения 0..511. Для VME это по сути псевдоадрес в диапазоне 256..511.
            ''' </summary>
            VI_ATTR_VXI_LA = 1073676501
            ''' <summary>
            ''' Идентификатор производителя. Может принимать значения 0..0xFFF.
            ''' </summary>
            VI_ATTR_MANF_ID = 1073676505
            ''' <summary>
            ''' Размер памяти, запрошенной устройством в адресном пространстве шины VXI.
            ''' </summary>
            VI_ATTR_MEM_SIZE = 1073676509
            ''' <summary>
            ''' Адресное пространство шины VXI. Значения из <see cref="Native.MemorySpace"/>.
            ''' </summary>
            VI_ATTR_MEM_SPACE
            ''' <summary>
            ''' Код модели устройства. Может принимать значения 0..0xFFFF.
            ''' </summary>
            VI_ATTR_MODEL_CODE
            ''' <summary>
            ''' Физический слот устройства VXI. Значения из <see cref="Native.Slot"/>.
            ''' </summary>
            VI_ATTR_SLOT = 1073676520
            ''' <summary>
            ''' Человекочитаемое текстовое описание интерфейса.
            ''' </summary>
            VI_ATTR_INTF_INST_NAME = -1073807127
            ''' <summary>
            ''' Задаёт, обслуживает ли устройство контроллер, на котором работает VISA. Значения <see cref="Native.VI_BOOL"/>.
            ''' </summary>
            VI_ATTR_IMMEDIATE_SERV = 1073676544
            ''' <summary>
            ''' Номер платы GPIB, к которой подключено устройство. Может принимать значения 0..0xFFFF.
            ''' </summary>
            VI_ATTR_INTF_PARENT_NUM
            ''' <summary>
            ''' Версия ресурса.
            ''' </summary>
            VI_ATTR_RSRC_SPEC_VERSION = 1073676656
            ''' <summary>
            ''' Тип интерфейса ресурса. Значения могут быть <see cref="InterfaceType"/>.
            ''' </summary>
            VI_ATTR_INTF_TYPE
            ''' <summary>
            ''' Первичный адрес локального GPIB контроллера в сессии (0..30).
            ''' </summary>
            VI_ATTR_GPIB_PRIMARY_ADDR
            ''' <summary>
            ''' Вторичный адрес локального GPIB контроллера в сессии. Значения <see cref="Native.GpibSecAddress"/>
            ''' </summary>
            VI_ATTR_GPIB_SECONDARY_ADDR
            ''' <summary>
            ''' Название производителя устройства.
            ''' </summary>
            VI_ATTR_RSRC_MANF_NAME = -1073806988
            ''' <summary>
            ''' Идентификатор производителя.
            ''' </summary>
            VI_ATTR_RSRC_MANF_ID = 1073676661
            ''' <summary>
            ''' Номер интерфейсной платы ресурса.
            ''' </summary>
            VI_ATTR_INTF_NUM
            ''' <summary>
            ''' Вид прерывания (аппаратное или программное <see cref="Native.TrigId.VI_TRIG_SW"/>). Значения могут быть <see cref="Native.TrigId"/>.
            ''' </summary>
            VI_ATTR_TRIG_ID
            ''' <summary>
            ''' Текущее состояние линии REN. Значения могут быть <see cref="Native.State"/>.
            ''' </summary>
            VI_ATTR_GPIB_REN_STATE = 1073676673
            ''' <summary>
            ''' Снимать ли адрес устройства (UNT и UNL) после каждой передачи данных. Значения <see cref="Native.VI_BOOL"/>.
            ''' </summary>
            VI_ATTR_GPIB_UNADDR_EN = 1073676676
            ''' <summary>
            ''' Задаёт байт статсуа в стиле 488 для локального контроллера в текущей сессии.
            ''' </summary>
            VI_ATTR_DEV_STATUS_BYTE = 1073676681
            ''' <summary>
            ''' При сохранении в файл добавлять ли данные или перезаписывать. Значения <see cref="Native.VI_BOOL"/>.
            ''' </summary>
            VI_ATTR_FILE_APPEND_EN = 1073676690
            ''' <summary>
            ''' Какие линии прерывания поддерживаются данным VXI. Битовый вектор, где биты 0..9 соответствуют <see cref="Native.TrigId.VI_TRIG_TTL0"/>..<see cref="Native.TrigId.VI_TRIG_ECL1"/>.
            ''' </summary>
            VI_ATTR_VXI_TRIG_SUPPORT = 1073676692
            ''' <summary>
            ''' TCP/IP адрес устройства в данной сессии. 
            ''' </summary>
            VI_ATTR_TCPIP_ADDR = -1073806955
            ''' <summary>
            ''' Имя хоста устройства.
            ''' </summary>
            VI_ATTR_TCPIP_HOSTNAME
            ''' <summary>
            ''' Задаёт TCP/IP порт.
            ''' </summary>
            VI_ATTR_TCPIP_PORT = 1073676695
            ''' <summary>
            ''' Имя устройства в локальной сети. 
            ''' </summary>
            VI_ATTR_TCPIP_DEVICE_NAME = -1073806951
            ''' <summary>
            ''' Когда этот атрибут задан, отключается алгоритм Нейгла. Атрибут по умолчанию включён.
            ''' Алгоритм Нейгла повышает скорость передачи в сети за счёт буферизации отправляемых данных до момента, пока не будет сформирован полноразмерный пакет.
            ''' </summary>
            VI_ATTR_TCPIP_NODELAY = 1073676698
            ''' <summary>
            ''' Атрибут позволяет приложению запросить включить пакеты "keep-alive". 
            ''' </summary>
            VI_ATTR_TCPIP_KEEPALIVE
            ''' <summary>
            ''' Задаёт, совместимо ли устройство с 488.2.
            ''' </summary>
            VI_ATTR_4882_COMPLIANT = 1073676703
            ''' <summary>
            ''' Серийный номер USB для устройства.
            ''' </summary>
            VI_ATTR_USB_SERIAL_NUM = -1073806944
            ''' <summary>
            ''' Номер интерфейса USB в текущей сессии.
            ''' </summary>
            VI_ATTR_USB_INTFC_NUM = 1073676705
            ''' <summary>
            ''' Протокол USB, используемый данным интерфейсом USB.
            ''' </summary>
            VI_ATTR_USB_PROTOCOL = 1073676711
            ''' <summary>
            ''' Максимальный размер данных, котрые будут сохраниться по любому прерыванию USB. Если данных больше, чем задано атрибутом, часть данных будет утеряна. 
            ''' Принимает значения 0..0xFFFF.
            ''' </summary>
            VI_ATTR_USB_MAX_INTR_SIZE = 1073676719
            ''' <summary>
            ''' Идентификатор выполняющейся команды.
            ''' </summary>
            VI_ATTR_JOB_ID = 1073692678
            ''' <summary>
            ''' Уникальный логический идентификатор события.
            ''' </summary>
            VI_ATTR_EVENT_TYPE = 1073692688
            ''' <summary>
            ''' Значение статуса/идентификатора, полученное в процессе цикла IACK или от регистра сигналов. Значения 0..FFFF.
            ''' </summary>
            VI_ATTR_SIGP_STATUS_ID
            ''' <summary>
            ''' Идентификатор механизма прерываний, по которому получено заданное прерывание.
            ''' Значения <see cref="Native.TrigId.VI_TRIG_TTL0"/>..<see cref="Native.TrigId.VI_TRIG_ECL1"/>.
            ''' </summary>
            VI_ATTR_RECV_TRIG_ID
            ''' <summary>
            ''' Значение статуса/идентификатора, полученное в процессе цикла IACK. Значения 0..FFFF_FFFF.
            ''' </summary>
            VI_ATTR_INTR_STATUS_ID = 1073692707
            ''' <summary>
            ''' Код завершённой асинхронной опреации ввода-вывода.
            ''' </summary>
            VI_ATTR_STATUS = 1073692709
            ''' <summary>
            ''' Действительный номер асинхронно переданных элементов.
            ''' </summary>
            VI_ATTR_RET_COUNT
            ''' <summary>
            ''' Адрес буфера, используемого в асинхронной операции.
            ''' </summary>
            VI_ATTR_BUFFER
            ''' <summary>
            ''' Уровень прерываний VXI, по которому было получено прерывание. Значения <see cref="Native.Levels"/>.
            ''' </summary>
            VI_ATTR_RECV_INTR_LEVEL = 1073692737
            ''' <summary>
            ''' Название операции, генерирующей событие.
            ''' </summary>
            VI_ATTR_OPER_NAME = -1073790910
            ''' <summary>
            ''' Контроллер перешёл в состояние controller-in-charge. Значения <see cref="Native.VI_BOOL.VI_TRUE"/>.
            ''' </summary>
            VI_ATTR_GPIB_RECV_CIC_STATE = 1073693075
            ''' <summary>
            ''' TCP/IP адрес устройства, от которого сессия получила соединение.
            ''' </summary>
            VI_ATTR_RECV_TCPIP_ADDR = -1073790568
            ''' <summary>
            ''' Число сохранённых байтов по прерыванию USB.
            ''' </summary>
            VI_ATTR_USB_RECV_INTR_SIZE = 1073693104
            ''' <summary>
            ''' Данные, полученные по прерыванию USB.
            ''' </summary>
            VI_ATTR_USB_RECV_INTR_DATA = -1073790543
            ''' <summary>
            ''' Время ожидания при поиске устройств VXI, мс (только для R&amp;S).
            ''' </summary>
            VI_RS_ATTR_TCPIP_FIND_RSRC_TMO = CInt(&HFAF0001UL)
            ''' <summary>
            ''' Режим поиска устройств VXI (только для R&amp;S). Значения из <see cref="RsFindModes"/>.
            ''' </summary>
            VI_RS_ATTR_TCPIP_FIND_RSRC_MODE = CInt(&HFAF0002UL)
            ''' <summary>
            ''' LXI производитель (только для R&amp;S).
            ''' </summary>
            VI_RS_ATTR_LXI_MANF = &H8FAF0003
            ''' <summary>
            ''' Модель LXI (только для R&amp;S).
            ''' </summary>
            VI_RS_ATTR_LXI_MODEL = &H8FAF0004
            ''' <summary>
            ''' Серийный номер LXI (только для R&amp;S).
            ''' </summary>
            VI_RS_ATTR_LXI_SERIAL = &H8FAF0005
            ''' <summary>
            ''' Версия прошивки LXI (только для R&amp;S).
            ''' </summary>
            VI_RS_ATTR_LXI_VERSION = &H8FAF0006
            ''' <summary>
            ''' Описание устройства LXI (только для R&amp;S).
            ''' </summary>
            VI_RS_ATTR_LXI_DESCRIPTION = &H8FAF0007
            ''' <summary>
            ''' Сетевое имя LXI устройства (только для R&amp;S).
            ''' </summary>
            VI_RS_ATTR_LXI_HOSTNAME = &H8FAF0008
        End Enum

        <Flags()>
        Public Enum RsFindModes As Integer
            ''' <summary>
            ''' Пропускать текущие (статические) сконфигурированные TCP/IP ресурсы.
            ''' </summary>
            VI_RS_FIND_MODE_NONE = 0
            ''' <summary>
            ''' Поиск статических сконфигурированных устройств LAN.
            ''' </summary>
            VI_RS_FIND_MODE_CONFIG = 1
            ''' <summary>
            ''' Искать все сетевые карты устройств VXI-11.
            ''' </summary>
            VI_RS_FIND_MODE_VXI11 = 2
            ''' <summary>
            ''' Искать все сетевые карты устройств LXI > v.1.3.
            ''' </summary>
            VI_RS_FIND_MODE_MDNS = 4
        End Enum

        ''' <summary>
        ''' Типы уведомлений.
        ''' </summary>
        Public Enum ViEventType As Integer
            ''' <summary>
            ''' Уведомление о завершении асинхронной операции.
            ''' </summary>
            VI_EVENT_IO_COMPLETION = 1073684489
            ''' <summary>
            ''' Уведомление, получено аппаратное прерывание.
            ''' </summary>
            VI_EVENT_TRIG = -1073799158
            ''' <summary>
            ''' Уведомление, что от оборудования пришёл запрос на обслуживание.
            ''' </summary>
            VI_EVENT_SERVICE_REQ = 1073684491
            ''' <summary>
            ''' Уведомление, что контроллер GPIB послал сообщение очистить устройство.
            ''' </summary>
            VI_EVENT_CLEAR = 1073684493
            ''' <summary>
            ''' Уведомление об ошибке в процессе вызова операции.
            ''' </summary>
            VI_EVENT_EXCEPTION = -1073799154
            ''' <summary>
            ''' Уведомление, что контроллер GPIB получил или потерял статус CIC (controller in charge).
            ''' </summary>
            VI_EVENT_GPIB_CIC = 1073684498
            ''' <summary>
            ''' Уведомление, что контроллер GPIB назначен передающим.
            ''' </summary>
            VI_EVENT_GPIB_TALK
            ''' <summary>
            ''' Уведомление, что контроллер GPIB назначен получающим.
            ''' </summary>
            VI_EVENT_GPIB_LISTEN
            ''' <summary>
            ''' Уведомление, что установлена линия SYSFAIL.
            ''' </summary>
            VI_EVENT_VXI_VME_SYSFAIL = 1073684509
            ''' <summary>
            ''' Уведомление, что линия SYSRESET сброшена.
            ''' </summary>
            VI_EVENT_VXI_VME_SYSRESET
            ''' <summary>
            ''' Уведомление, что от оборудования пришёл сигнал VXI или прерывание VXI.
            ''' </summary>
            VI_EVENT_VXI_SIGP = 1073684512
            ''' <summary>
            ''' Уведомление, что от устройства получено прерывание шины VXI.
            ''' </summary>
            VI_EVENT_VXI_VME_INTR = -1073799135
            ''' <summary>
            ''' Уведомление, что установлено TCP/IP соединение.
            ''' </summary>
            VI_EVENT_TCPIP_CONNECT = 1073684534
            ''' <summary>
            ''' Уведомление о событии USB.
            ''' </summary>
            VI_EVENT_USB_INTR
            ''' <summary>
            ''' Все активные уведомления.
            ''' </summary>
            VI_ALL_ENABLED_EVENTS = 1073709055
        End Enum

        ''' <summary>
        ''' Коды статуса завершения операций.
        ''' </summary> 
        Public Enum ViStatus As Integer
            ''' <summary>
            ''' Операция выполнена успешно.
            ''' </summary>
            VI_SUCCESS = 0
            ''' <summary>
            ''' Событие установлено как минимум для одного из заданных механизмов.
            ''' </summary>
            VI_SUCCESS_EVENT_EN = 1073676290
            ''' <summary>
            ''' Заданное событие успешно отключено.
            ''' </summary>
            VI_SUCCESS_EVENT_DIS = 1073676291
            ''' <summary>
            ''' Очередь событий успешно очищена.
            ''' </summary>
            VI_SUCCESS_QUEUE_EMPTY = 1073676292
            ''' <summary>
            ''' Прочитан заданный символ разделитель.
            ''' </summary>
            VI_SUCCESS_TERM_CHAR = 1073676293
            ''' <summary>
            ''' Число прочитанных байтов равно ожидаемому.
            ''' </summary>
            VI_SUCCESS_MAX_CNT = 1073676294
            ''' <summary>
            ''' Сессия открыта успешно, но устройство по специфическому адресу не отвечает.
            ''' </summary>
            VI_SUCCESS_DEV_NPRESENT = 1073676413
            ''' <summary>
            ''' Путь от источника до получателя триггера успешно назначен.
            ''' </summary>
            VI_SUCCESS_TRIG_MAPPED = 1073676414
            ''' <summary>
            ''' Ожидание успешно снято при получении уведомления о событии. Ещё есть хотя бы одно доступное событие для данной сессии.
            ''' </summary>
            VI_SUCCESS_QUEUE_NEMPTY = 1073676416
            ''' <summary>
            ''' Событие обработано успешно. Не вызывайте другие обработчики данного события в данной сессии.
            ''' </summary>
            VI_SUCCESS_NCHAIN = 1073676440
            ''' <summary>
            ''' Заданный режим доступа успешно применён и сессия получила внутренние общие блокировки.
            ''' </summary>
            VI_SUCCESS_NESTED_SHARED = 1073676441
            ''' <summary>
            ''' Заданный режим доступа успешно применён и сессия получила внутренние эксклюзивные блокировки.
            ''' </summary>
            VI_SUCCESS_NESTED_EXCLUSIVE = 1073676442
            ''' <summary>
            ''' Операция чтения или записи выполнена синхронно.
            ''' </summary>
            VI_SUCCESS_SYNC = 1073676443
            ''' <summary>
            ''' Возвращённое значение корректно. Одно или более событий не возникли по причине нехватки доступного места на время их возникновения.
            ''' </summary>
            VI_WARN_QUEUE_OVERFLOW = 1073676300
            ''' <summary>
            ''' По крайней мере один модуль не может быть загружен.
            ''' </summary>
            VI_WARN_CONFIG_NLOADED = 1073676407
            ''' <summary>
            ''' Ссылка на переданный объект недействительна.
            ''' </summary>
            VI_WARN_NULL_OBJECT = 1073676418
            ''' <summary>
            ''' Хотя значение атрибута валидное, оно не поддерживается реализацией ресурса.
            ''' </summary>
            VI_WARN_NSUP_ATTR_STATE = 1073676420
            ''' <summary>
            ''' Код состояния, переданный функции, не распознан.
            ''' </summary>
            VI_WARN_UNKNOWN_STATUS = 1073676421
            ''' <summary>
            ''' Заданный буфер не поддерживается.
            ''' </summary>
            VI_WARN_NSUP_BUF = 1073676424
            ''' <summary>
            ''' Операция выполнена, но в драйвере нижнего уровня не реализована расширенная функциональность.
            ''' </summary>
            VI_WARN_EXT_FUNC_NIMPL = 1073676457
            ''' <summary>
            ''' Сбой инициализации VISA.
            ''' </summary>
            VI_ERROR_SYSTEM_ERROR = -1073807360
            ''' <summary>
            ''' Заданный идентификатор сессии недействителен.
            ''' </summary>
            VI_ERROR_INV_OBJECT = -1073807346
            ''' <summary>
            ''' Заданный тип блокировки не может быть получен/установлен т.к. ресурс уже заблокирован несовместимым с запрашиваемым режимом блокировки.
            ''' </summary>
            VI_ERROR_RSRC_LOCKED = -1073807345
            ''' <summary>
            ''' Недействительное выражение для поиска ресурса.
            ''' </summary>
            VI_ERROR_INV_EXPR = -1073807344
            ''' <summary>
            ''' Неверная информация о ресурсе или ресурс не представлен в системе.
            ''' </summary>
            VI_ERROR_RSRC_NFOUND = -1073807343
            ''' <summary>
            ''' Недействительный идентификатор ресурса. Ошибка парсера.
            ''' </summary>
            VI_ERROR_INV_RSRC_NAME = -1073807342
            ''' <summary>
            ''' Неверный режим доступа.
            ''' </summary>
            VI_ERROR_INV_ACC_MODE = -1073807341
            ''' <summary>
            ''' Истёк заданный период времени (таймаут).
            ''' </summary>
            VI_ERROR_TMO = -1073807339
            ''' <summary>
            ''' Невозможно деаллоцировать ранее выделенные структуры данных по переданному идентификатору.
            ''' </summary>
            VI_ERROR_CLOSING_FAILED = -1073807338
            ''' <summary>
            ''' Задано недействительное значение.
            ''' </summary>
            VI_ERROR_INV_DEGREE = -1073807333
            ''' <summary>
            ''' Задан недействительный идентификатор задания.
            ''' </summary>
            VI_ERROR_INV_JOB_ID = -1073807332
            ''' <summary>
            ''' Атрибут не определён в заданном ресурсе.
            ''' </summary>
            VI_ERROR_NSUP_ATTR = -1073807331
            ''' <summary>
            ''' Заданное значение атрибута недействительно или не поддерживается ресурсом.
            ''' </summary>
            VI_ERROR_NSUP_ATTR_STATE = -1073807330
            ''' <summary>
            ''' Заданный атрибут только для чтения.
            ''' </summary>
            VI_ERROR_ATTR_READONLY = -1073807329
            ''' <summary>
            ''' Заданный тип блокировки не поддерживается ресурсом.
            ''' </summary>
            VI_ERROR_INV_LOCK_TYPE = -1073807328
            ''' <summary>
            ''' Переданное значение requestedKey не является допустимым ключом доступа к заданному ресурсу.
            ''' </summary>
            VI_ERROR_INV_ACCESS_KEY = -1073807327
            ''' <summary>
            ''' Тип события недопустим для заданного ресурса.
            ''' </summary>
            VI_ERROR_INV_EVENT = -1073807322
            ''' <summary>
            ''' Механизм, заданный для события, недействительный.
            ''' </summary>
            VI_ERROR_INV_MECH = -1073807321
            ''' <summary>
            ''' Обработчик не назначен для заданного события. Сессия не может быть создана для режима VI_HNDLR механизма обратного вызова.
            ''' </summary>
            VI_ERROR_HNDLR_NINSTALLED = -1073807320
            ''' <summary>
            ''' Заданный указатель на обработчик недействительный.
            ''' </summary>
            VI_ERROR_INV_HNDLR_REF = -1073807319
            ''' <summary>
            ''' Заданный контекст события недействительный. 
            ''' </summary>
            VI_ERROR_INV_CONTEXT = -1073807318
            ''' <summary>
            ''' Чтобы получать события заданного типа, должна быть установлена сессия.
            ''' </summary>
            VI_ERROR_NENABLED = -1073807313
            ''' <summary>
            ''' Вызовы во время исполнения для заданного ресурса прерваны.
            ''' </summary>
            VI_ERROR_ABORT = -1073807312
            ''' <summary>
            ''' В процессе передачи произошло нарушение протокола записи.
            ''' </summary>
            VI_ERROR_RAW_WR_PROT_VIOL = -1073807308
            ''' <summary>
            ''' В процессе передачи произошло нарушение протокола чтения.
            ''' </summary>
            VI_ERROR_RAW_RD_PROT_VIOL = -1073807307
            ''' <summary>
            ''' В процессе передачи произошла ошибка выходного протокола.
            ''' </summary>
            VI_ERROR_OUTP_PROT_VIOL = -1073807306
            ''' <summary>
            ''' В процессе передачи произошла ошибка входного протокола.
            ''' </summary>
            VI_ERROR_INP_PROT_VIOL = -1073807305
            ''' <summary>
            ''' В процессе передачи произошла ошибка шины.
            ''' </summary>
            VI_ERROR_BERR = -1073807304
            ''' <summary>
            ''' Невозможно запросить асинхронную операцию по причине того, что операция уже выполняется.
            ''' </summary>
            VI_ERROR_IN_PROGRESS = -1073807303
            ''' <summary>
            ''' Некоторые специфические конфигурации не существуют или не могут быть загружены.
            ''' </summary>
            VI_ERROR_INV_SETUP = -1073807302
            ''' <summary>
            ''' Невозможно запросить операцию чтения.
            ''' </summary>
            VI_ERROR_QUEUE_ERROR = -1073807301
            ''' <summary>
            ''' Недостаточно системных ресурсов для создания сессии.
            ''' </summary>
            VI_ERROR_ALLOC = -1073807300
            ''' <summary>
            ''' Система не может задать буфер для заданной маски.
            ''' </summary>
            VI_ERROR_INV_MASK = -1073807299
            ''' <summary>
            ''' Неизвестная ошибка ввода-вывода в процессе передачи.
            ''' </summary>
            VI_ERROR_IO = -1073807298
            ''' <summary>
            ''' Строка, задающая формат в writeFmt(), неверна.
            ''' </summary>
            VI_ERROR_INV_FMT = -1073807297
            ''' <summary>
            ''' Строка, задающая формат в writeFmt(), не поддерживается.
            ''' </summary>
            VI_ERROR_NSUP_FMT = -1073807295
            ''' <summary>
            ''' Строка триггера (trigSrc или trigDest) сейчас используется.
            ''' </summary>
            VI_ERROR_LINE_IN_USE = -1073807294
            ''' <summary>
            ''' Заданный режим не поддерживается реализацией VISA.
            ''' </summary>
            VI_ERROR_NSUP_MODE = -1073807290
            ''' <summary>
            ''' Запрос на обслуживание не получен для сессии.
            ''' </summary>
            VI_ERROR_SRQ_NOCCURRED = -1073807286
            ''' <summary>
            ''' Задано недействительное адресное пространство.
            ''' </summary>
            VI_ERROR_INV_SPACE = -1073807282
            ''' <summary>
            ''' Заданное смещение недействительно.
            ''' </summary>
            VI_ERROR_INV_OFFSET = -1073807279
            ''' <summary>
            ''' Задано недействительная длина источника или получателя.
            ''' </summary>
            VI_ERROR_INV_WIDTH = -1073807278
            ''' <summary>
            ''' Недействительное смещение источника или получателя недоступно на данном оборудовании.
            ''' </summary>
            VI_ERROR_NSUP_OFFSET = -1073807276
            ''' <summary>
            ''' Невозможно поддержать различающиеся длины источника и получателя.
            ''' </summary>
            VI_ERROR_NSUP_VAR_WIDTH = -1073807275
            ''' <summary>
            ''' Сессия не назначена.
            ''' </summary>
            VI_ERROR_WINDOW_NMAPPED = -1073807273
            ''' <summary>
            ''' Предыдущий запрос в состоянии ожидания, вызванного ошибкой множественного запроса.
            ''' </summary>
            VI_ERROR_RESP_PENDING = -1073807271
            ''' <summary>
            ''' Не обнаружено слушателей (и NRFD, и NDAC).
            ''' </summary>
            VI_ERROR_NLISTENERS = -1073807265
            ''' <summary>
            ''' Связанный с ресурсом интерфейс не загружен.
            ''' </summary>
            VI_ERROR_NCIC = -1073807264
            ''' <summary>
            ''' Назначенный сессии интерфейс не является системным контроллером.
            ''' </summary>
            VI_ERROR_NSYS_CNTLR = -1073807263
            ''' <summary>
            ''' Заданная сессия не поддерживает операцию. Сессия поддерживается только менеджером ресурсов.
            ''' </summary>
            VI_ERROR_NSUP_OPER = -1073807257
            ''' <summary>
            ''' Прерывание в процессе ожидания от предыдущего вызова.
            ''' </summary>
            VI_ERROR_INTR_PENDING = -1073807256
            ''' <summary>
            ''' Ошибка проверки чётности в процессе передачи.
            ''' </summary>
            VI_ERROR_ASRL_PARITY = -1073807254
            ''' <summary>
            ''' Ошибка кадров в процессе передачи.
            ''' </summary>
            VI_ERROR_ASRL_FRAMING = -1073807253
            ''' <summary>
            ''' Ошибка набегания в процессе передачи. Символ не прочитан до того, как от устройства пришёл предыдущий символ.
            ''' </summary>
            VI_ERROR_ASRL_OVERRUN = -1073807252
            ''' <summary>
            ''' Путь от trigSrc до trigDest не назначен.
            ''' </summary>
            VI_ERROR_TRIG_NMAPPED = -1073807250
            ''' <summary>
            ''' Заданное смещение некорректно выровнено.
            ''' </summary>
            VI_ERROR_NSUP_ALIGN_OFFSET = -1073807248
            ''' <summary>
            ''' Заданный пользовательский буфер требуемого размера недействителен или недоступен.
            ''' </summary>
            VI_ERROR_USER_BUF = -1073807247
            ''' <summary>
            ''' Ресурс действительный, но VISA сейчас не может получить к нему доступ.
            ''' </summary>
            VI_ERROR_RSRC_BUSY = -1073807246
            ''' <summary>
            ''' Заданная ширина не поддерживается оборудованием.
            ''' </summary>
            VI_ERROR_NSUP_WIDTH = -1073807242
            ''' <summary>
            ''' Значение параметра некорректно.
            ''' </summary>
            VI_ERROR_INV_PARAMETER = -1073807240
            ''' <summary>
            ''' Заданный протокол некорректен.
            ''' </summary>
            VI_ERROR_INV_PROT = -1073807239
            ''' <summary>
            ''' Заданный размер некорректен.
            ''' </summary>
            VI_ERROR_INV_SIZE = -1073807237
            ''' <summary>
            ''' Заданная сессия уже содержит назначенное окно.
            ''' </summary>
            VI_ERROR_WINDOW_MAPPED = -1073807232
            ''' <summary>
            ''' Операция не реализована.
            ''' </summary>
            VI_ERROR_NIMPL_OPER = -1073807231
            ''' <summary>
            ''' Задана недействительаная длина.
            ''' </summary>
            VI_ERROR_INV_LENGTH = -1073807229
            ''' <summary>
            ''' Значение режима недействительно.
            ''' </summary>
            VI_ERROR_INV_MODE = -1073807215
            ''' <summary>
            ''' Текущая сессия не имеет блокировок ресурса.
            ''' </summary>
            VI_ERROR_SESN_NLOCKED = -1073807204
            ''' <summary>
            ''' Устройство не экспортирует память.
            ''' </summary>
            VI_ERROR_MEM_NSHARED = -1073807203
            ''' <summary>
            ''' Библиотека не может быть найдена или загружена.
            ''' </summary>
            VI_ERROR_LIBRARY_NFOUND = -1073807202
            ''' <summary>
            ''' Интерфейс не может сгенерировать прерывание по заданному уровню или с запрошенным значением statusID.
            ''' </summary>
            VI_ERROR_NSUP_INTR = -1073807201
            ''' <summary>
            ''' Заданое значение линейного параметра недействительно.
            ''' </summary>
            VI_ERROR_INV_LINE = -1073807200
            ''' <summary>
            ''' Ошибка открытия заданного файла. Возможные причины - недействительный путь или отсутствие разрешения на доступ.
            ''' </summary>
            VI_ERROR_FILE_ACCESS = -1073807199
            ''' <summary>
            ''' Ошибка в процессе доступа к заданному файлу.
            ''' </summary>
            VI_ERROR_FILE_IO = -1073807198
            ''' <summary>
            ''' Одна из линий (trigSrc или trigDest) не поддерживается данной реализацией VISA.
            ''' </summary>
            VI_ERROR_NSUP_LINE = -1073807197
            ''' <summary>
            ''' Механизм не поддерживается данным типом события.
            ''' </summary>
            VI_ERROR_NSUP_MECH = -1073807196
            ''' <summary>
            ''' Тип интерфейса действительный, но он не сконфигурирован.
            ''' </summary>
            VI_ERROR_INTF_NUM_NCONFIG = -1073807195
            ''' <summary>
            ''' Соединение в данной сессии потеряно.
            ''' </summary>
            VI_ERROR_CONN_LOST = -1073807194
        End Enum

        Public Enum ManufacturerIds As UShort
            None = 0
            ''' <summary>
            ''' В директории system32 стандартная библиотека по умолчанию visa32.dll.
            ''' </summary>
            MANFID_DEFAULT = UShort.MaxValue
            ''' <summary>
            ''' В директории system32 библиотека rsvisa32.dll.
            ''' </summary>
            MANFID_RS = 4015US
            ''' <summary>
            ''' В директории system32 библиотека nivisa32.dll.
            ''' </summary>
            MANFID_NI = 4086US
            ''' <summary>
            ''' В директории system32 библиотека agvisa32.dll.
            ''' </summary>
            MANFID_AG = 4095US
        End Enum

        ''' <summary>
        ''' Типы интерфейса.
        ''' </summary>
        Public Enum InterfaceType As Short
            None = 0
            VI_INTF_GPIB = 1S
            VI_INTF_VXI = 2S
            VI_INTF_GPIB_VXI = 3S
            VI_INTF_ASRL = 4S
            VI_INTF_TCPIP = 6S
            VI_INTF_USB = 7S
        End Enum

#End Region '/ENUMS

#End Region '/NESTED TYPES

    End Class '/VisaDevice

End Namespace