# visa_vb
Оболочка для работы с **VISA** (*Virtual Instrument Software Architecture*) на Visual Basic .NET.

## Описание

Открытый интерфейс класса *VisaDevice.vb* содержит минимально необходимую функциональность для работы с измерительной аппаратурой по протоколу VISA. 

Во вложенном классе *Native* содержатся объявления всех доступных в библиотеке `visa32.dll` функций, которые при необходимости можно с лёгкостью задействовать в своём проекте.

### Примеры использования

- Отображение идентификатора прибора:

```
Dim devs As String() = VisaDevice.GetVisaDevices()

Using vi As New VisaDevice(devs(0))
  Console.WriteLine($"IDN = {vi.ShowIdn()}")
End Using
```

- Запрос мощности у измерителя мощности NRP:

```
''' <summary>
''' Запрашивает у NRP и возвращает измеренную мощность, в дБм.
''' </summary>
''' <param name="sensorNumber">Номер датчика.</param>
Public Function GetPower(Optional sensorNumber As Integer = 1) As Double
  SendQuery("init:cont 0")
  SendQuery("init:imm")
  Dim res As String
  Dim cnt As Integer = 0
  Do
    res = SendQuery("*opc?")
    cnt += 1
  Loop Until (res.Trim() = "1") OrElse (cnt > 20)
  Dim ans As String = SendQuery($"fetch{sensorNumber}:scalar:pow?")
  SendQuery("init:cont 1")
  Dim power As Double = Double.Parse(ans, New Globalization.CultureInfo("en-US"))
  Return power
End Function
```
