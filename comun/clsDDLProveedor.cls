VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia TABLA
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get table() As String

    table = "proveedor"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorID() As String

    proveedorID = "proveedor_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorID_t() As String

    proveedorID_t = table & ".proveedor_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia razonSocial
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get razonSocial() As String

    razonSocial = "razon_social"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia razonSocial completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get razonSocial_t() As String

    razonSocial_t = table & ".razon_social"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia nombreFantasia
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombreFantasia() As String

    nombreFantasia = "nombre_fantasia"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia nombreFantasia completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombreFantasia_t() As String

    nombreFantasia_t = table & ".nombre_fantasia"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia cuit
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuit() As String

    cuit = "cuit"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia cuit completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuit_t() As String

    cuit_t = table & ".cuit"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia domicilio
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get domicilio() As String

    domicilio = "domicilio"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia domicilio completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get domicilio_t() As String

    domicilio_t = table & ".domicilio"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia localidad
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get localidad() As String

    localidad = "localidad"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia localidad completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get localidad_t() As String

    localidad_t = table & ".localidad"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia provincia
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get provincia() As String

    provincia = "provincia"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia provincia completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get provincia_t() As String

    provincia_t = table & ".provincia"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia telefono
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get telefono() As String

    telefono = "telefono"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia telefono completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get telefono_t() As String

    telefono_t = table & ".telefono"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fax
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fax() As String

    fax = "fax"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fax completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fax_t() As String

    fax_t = table & ".fax"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia email
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get email() As String

    email = "email"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia email completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get email_t() As String

    email_t = table & ".email"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia posicionIva
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get posicionIva() As String

    posicionIva = "posicion_iva"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia posicionIva completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get posicionIva_t() As String

    posicionIva_t = table & ".posicion_iva"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia celular
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get celular() As String

    celular = "celular"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia celular completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get celular_t() As String

    celular_t = table & ".celular"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia ingresosBrutos
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ingresosBrutos() As String

    ingresosBrutos = "ingresos_brutos"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia ingresosBrutos completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ingresosBrutos_t() As String

    ingresosBrutos_t = table & ".ingresos_brutos"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia contacto
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get contacto() As String

    contacto = "contacto"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia contacto completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get contacto_t() As String

    contacto_t = table & ".contacto"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia observaciones
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get observaciones() As String

    observaciones = "observaciones"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia observaciones completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get observaciones_t() As String

    observaciones_t = table & ".observaciones"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia autoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As String

    autoID = "auto_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia autoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID_t() As String

    autoID_t = table & ".auto_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia created
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As String

    created = "created"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia created completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created_t() As String

    created_t = table & ".created"

End Property

