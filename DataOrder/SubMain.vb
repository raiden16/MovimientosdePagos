Option Strict Off
Option Explicit On

Module SubMain

    Friend oCatchingEvents As CatchingEvents

    Sub Main()

        oCatchingEvents = New CatchingEvents
        ' oCatchingEvents.comprobarlicencia()
        System.Windows.Forms.Application.Run()

    End Sub

End Module
