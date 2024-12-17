Module Modulovars

    

    Public imagenlogocea As Image
    'email
    Public email_tipodes, email_des, email_maildes, email_asunto, email_file, email_vienede, email_message As String

    Public VIENEDE_dto As String
    Public VIENEDE_alumno As String = "no"
    Public VIENEDE_alumno_ok As String = "no"


    Public num_nomina_global As String

    'DATOS DE LA CAADEMIA
    Public SQLSOURCE As String
    Public aca_nom, aca_nom2, aca_nit, aca_tels, aca_cels, aca_dirs, aca_mail, aca_prop, aca_regimen, aca_lic_min, aca_lic_sec, aca_web, dian_res, dian_fecha, dian_rango As String
    Public aca_logoname As String

    'vars de usuarios del sistema
    Public usrdoc As Double
    Public usrnom, usrpass, usrtipo, usrusr As String

    Public mycurso As Long
    Public ultimafact As Long
    Public ultimo_recibo As Long
    Public ultimo_recibo_egreso, ultimo_recibo_egreso_nomina, ultimo_recibo_ingreso As Long
    Public ultimo_cargue As Long

    'vars de tramitador
    Public tramdoc As Double
    Public tramnom, tramtel As String

    Public tipohoraglobal As String

    'vars alumnos
    Public aldoc, altipodoc, alrunt, alnom1, alnom2, alape1, alape2, algenero, aldir, alresidencia, alorigen, alfechan, altel, alcel, almail As String
    Public alestadoc, aleduca, alocupa, alestrato, alsalud, aldiscapa, almulticul, ALMRH As String


    'vars imprimir factura
    Public imp_M As String = "NO"

    Public imp_D As String = "NO"

    Public imp_nuevosaldo As String = "0"
    Public imp_factunum As Integer
    Public imp_clientedoc, imp_clientenom, imp_clientetel, imp_clientedir, imp_fecha, imp_usuario, imp_estado As String
    Public imp_concepto, imp_concepto2, imp_asesor, imp_curso As String
    Public imp_valor, imp_subtotal, imp_total As Long
    Public imp_asesor_nom, imp_asesor_doc, imp_asesor_dir, imp_asesor_cel
    Public imp_prefijo_prioridad As String = "NO"
    Public imp_mediopago, imp_empleado As String
    Public imp_egre_fecha, imp_egre_clienom, imp_egre_cliedoc, imp_egrefpago, imp_egre_cliedir, imp_egre_clietel, imp_egre_concep, imp_egre_valor, imp_egre_empleado, imp_egre_estado As String
    Public imp_egre_num, imp_egre_concepvr, imp_egre_subttl, imp_egre_ttl As Long


    Public RC_VER As String = ""
    Public RC_VER_curso As String = ""
    Public RC_VER_AlumnoDoc As String = ""

    Public RC2_VER As String = ""


    Public CE_VER As String = ""

    Public CE2_VER As String = ""


End Module
