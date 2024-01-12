def read_breaker_switch(self, oSht):
    # read the shunt/filter data which may be required in the ipsa modelling.
    for row in range(2, self.iBreakerSwitch + 2):
        oDIgSILENT_BreakerSwitch = DIgSILENT_BreakerSwitch()
        oDIgSILENT_BreakerSwitch.sName = str(oSht.Cells(row, 1).Value)
        oDIgSILENT_BreakerSwitch.sTerminaliSubstation = str(oSht.Cells(row, 5).Value)
        oDIgSILENT_BreakerSwitch.sTerminali = str(oSht.Cells(row, 6).Value)
        oDIgSILENT_BreakerSwitch.sTerminali = self.delete_space_at_the_beginning(
            oDIgSILENT_BreakerSwitch.sTerminali)
        if oDIgSILENT_BreakerSwitch.sTerminali in self.list_sDuplicateTerminalName and self.bTerminalIDOnly:
            oDIgSILENT_BreakerSwitch.sTerminali = self.update_terminal_name(oDIgSILENT_BreakerSwitch.sTerminali,
                                                                            oDIgSILENT_BreakerSwitch.sTerminaliSubstation)
        else:
            pass
        oDIgSILENT_BreakerSwitch.sTerminali = self.delete_line_at_breaker_terminal(
            oDIgSILENT_BreakerSwitch.sTerminali)
        oDIgSILENT_BreakerSwitch.sTerminaljSubstation = str(oSht.Cells(row, 7).Value)
        oDIgSILENT_BreakerSwitch.sTerminalj = str(oSht.Cells(row, 8).Value)
        oDIgSILENT_BreakerSwitch.sTerminalj = self.delete_space_at_the_beginning(
            oDIgSILENT_BreakerSwitch.sTerminalj)
        if oDIgSILENT_BreakerSwitch.sTerminalj in self.list_sDuplicateTerminalName and self.bTerminalIDOnly:
            oDIgSILENT_BreakerSwitch.sTerminalj = self.update_terminal_name(oDIgSILENT_BreakerSwitch.sTerminalj,
                                                                            oDIgSILENT_BreakerSwitch.sTerminaljSubstation)
        else:
            pass
        oDIgSILENT_BreakerSwitch.sTerminalj = self.delete_line_at_breaker_terminal(
            oDIgSILENT_BreakerSwitch.sTerminalj)

        self.list_sUsedTerminalName.append(oDIgSILENT_BreakerSwitch.sTerminali)
        self.list_sUsedTerminalName.append(oDIgSILENT_BreakerSwitch.sTerminalj)

        oDIgSILENT_BreakerSwitch.iClosed = int(oSht.Cells(row, 12).Value)
        oDIgSILENT_BreakerSwitch.sActualState = oSht.Cells(row, 13).Value

        self.list_oBreakerSwitch.append(oDIgSILENT_BreakerSwitch)

    return