Public user as String
Public senha as String


Sub Organizar_tudo()

    Call organizar_SE16N
    Call organizar_ME5A
    Call organizar_JDE
    Call Pedidos_SAP
    Call Pedidos_JDE
    Call mesclar_Reqs



    Importar_Follow_JDE(dt_ini, dt_fin)
End Sub
