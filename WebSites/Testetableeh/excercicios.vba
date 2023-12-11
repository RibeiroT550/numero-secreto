sub excercicios()

dim valor1, valor2 as integer


valor1 = int((1000 - -1000 + 1) * rnd + 0)

valor2 = int((1000 - -1000 + 1) * rnd + 0)

valor = 0


if valor1> valor2 then

    ccu#Region "Module Description"

     valor = valor1

else

valor = valor2

end if

end if

if valor > 0 then
formatcurrency(valor)
            msgbox "O valor " & valor & " é positivo"     
    else
    msgbox "Nenhum valor é positivo"
    end if

    


end sub
