Sub Задание1()

Dim Ребро_Куба As Integer
Dim Обьем_Куба As Integer
Dim Площадь_Куба As Integer

Ребро_Куба = InputBox("Задайте длину ребра куба")
Обьем_Куба = Ребро_Куба ^ 3
Площадь_Куба = 6 * Ребро_Куба ^ 2

 MsgBox "Объем куба равен: " & Обьем_Куба & vbNewLine & _
           "Площадь куба равна: " & Площадь_Куба
           
End Sub

Sub Задание2()

Dim Сырье As Double
Dim Пряжа As Double
Dim Отходы As Double
Dim Потери As Double
    
Сырье = 12 * 1000
    
Пряжа = Сырье * 0.93
Отходы = Сырье * 0.06
Потери = Сырье * 0.01
    
MsgBox "Из 12 т сырья получим:" & vbNewLine & _
        "Пряжи: " & Пряжа & " кг" & vbNewLine & _
        "Отходов: " & Отходы & " кг" & vbNewLine & _
        "Потерь: " & Потери & " кг"
        
End Sub

Sub Задание3()

    Dim Длина As Double
    Dim Ширина As Double
    Dim Радиус As Double
    Dim Диаметр As Double
    Dim Количество As Integer
    Dim Площадь_Заготовки As Double
    Dim Площадь_Материи As Double
    Dim Площадь_Отходов As Double
    Dim ПИ  As Double
    
    Длина = 12
    Ширина = 1.4
    Радиус = 0.15
    Диаметр = Радиус * 2
    ПИ = 3.14
    
    Количество = Int(Длина / Диаметр)
    
    Площадь_Заготовки = ПИ * (Радиус ^ 2)
    
    Площадь_Материи = Длина * Ширина
    
    Площадь_Отходов = Площадь_Материи - Количество * Площадь_Заготовки
    
    MsgBox "Количество заготовок: " & Количество & vbNewLine & _
           "Площадь отходов: " & Площадь_Отходов
           
End Sub

Sub Задание4()

    Dim X1 As Double, Y1 As Double
    Dim X2 As Double, Y2 As Double
    Dim Расстояние As Double
    
    X1 = InputBox("Введите X1:")
    Y1 = InputBox("Введите Y1:")
    X2 = InputBox("Введите X2:")
    Y2 = InputBox("Введите Y2:")
    
    Расстояние = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
    
    MsgBox "Расстояние между точками: " & Расстояние
    
End Sub

Sub Задание6()

    Dim Производительность_одной As Double
    Dim Количество_машин As Integer
    Dim Время_часы As Double
    Dim Время_минуты As Double
    Dim Всего_тарелок As Double
    
    Производительность_одной = 7
    Количество_машин = 3
    Время_часы = 6
    Время_минуты = Время_часы * 60
    
    Всего_тарелок = Производительность_одной * Количество_машин * Время_минуты
    
    MsgBox "Три машины за 6 часов выпустят: " & Всего_тарелок & " тарелок"
    
End Sub

Sub Задание5()
    
    Dim A As Double, B As Double, C As Double
    Dim Сумма As Double
    Dim Процент_A As Double, Процент_B As Double, Процент_C As Double
    
    A = InputBox("Введите A:")
    B = InputBox("Введите B:")
    C = InputBox("Введите C:")
    
    Сумма = A + B - C
    
   
    
    Процент_A = (A / Сумма) * 100
    Процент_B = (B / Сумма) * 100
    Процент_C = (C / Сумма) * 100
    
    MsgBox "От (A + B - C) = " & Сумма & vbNewLine & _
           "На A приходится: " & Процент_A & "%" & vbNewLine & _
           "На B приходится: " & Процент_B & "%" & vbNewLine & _
           "На C приходится: " & Процент_C & "%"
End Sub
