object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 379
  Width = 437
  object FDConnection1: TFDConnection
    Params.Strings = (
      'DriverID=SQLite')
    Connected = True
    LoginPrompt = False
    Left = 72
    Top = 48
  end
  object FDPhysSQLiteDriverLink1: TFDPhysSQLiteDriverLink
    Left = 176
    Top = 64
  end
  object FDGUIxWaitCursor1: TFDGUIxWaitCursor
    Provider = 'Forms'
    Left = 264
    Top = 40
  end
  object FDQuery1: TFDQuery
    Connection = FDConnection1
    SQL.Strings = (
      'select * from contact')
    Left = 56
    Top = 152
  end
  object DataSource1: TDataSource
    DataSet = FDQuery1
    Left = 128
    Top = 160
  end
end
