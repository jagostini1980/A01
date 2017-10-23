if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOcOrdenesDeCompraRecepcionTraer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOcOrdenesDeCompraRecepcionTraer]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE SpOcOrdenesDeCompraRecepcionTraer 
					@P_Codigo  int
AS

	Select O_CodigoArticulo,IsNull(C_TablaArticulos,'')C_TablaArticulos , 
	           Sum(O_CantidadPendiente)O_CantidadPendiente, Sum(O_CantidadPedida) O_CantidadPedida , O_Autorizado
	   From OC_OrdenesDeCompraRenglones Reng
	   Join TA_CentrosDeCostos
	   On C_Codigo = O_CentroDeCostoEmisor 
	   Join
	  (Select O_NumeroOrdenDeCompra, O_CentroDeCostoEmisor, O_Autorizado
	   From OC_OrdenesDeCompraCabecera C
	   Where O_CodigoProveedor = @P_Codigo and O_FechaAnulacion is null
	   and Exists(Select O_NumeroOrdenDeCompra, O_CentroDeCostoEmisor
		      From OC_OrdenesDeCompraRenglones R
		      Where O_CantidadPendiente >0
			   And C.O_NumeroOrdenDeCompra = R.O_NumeroOrdenDeCompra 
			   And C.O_CentroDeCostoEmisor = R.O_CentroDeCostoEmisor)) Cab
	   On Cab.O_NumeroOrdenDeCompra = Reng.O_NumeroOrdenDeCompra 
	        and Cab.O_CentroDeCostoEmisor = Reng.O_CentroDeCostoEmisor
	   Group By O_CodigoArticulo, C_TablaArticulos,O_Autorizado
	HAVING SUM(O_CantidadPendiente) > 0

	
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeCompraCabeceraAgregar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeCompraCabeceraAgregar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeCompraCabeceraModificar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeCompraCabeceraModificar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeCompraCabeceraTraer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeCompraCabeceraTraer]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE SpOCOrdenesDeCompraCabeceraAgregar
	 @O_Fecha 				smalldatetime,
	 @O_Responsable 			varchar(50),
	 @O_CodigoProveedor 			int,
	 @O_LugarDeEntrega 			varchar(50),
	 @O_FormaDePagoPactada 		varchar(50),
	 @O_EmpresaFacturaANombreDe 	char(3),
	 @U_Usuario 				varchar(15),
	 @O_CentroDeCostoEmisor 		varchar(4),
	 @O_Observaciones			varchar(100),
	 @O_FechaEmision			DateTime,
	 @O_CodigoLugarDeEntrega		Int,
	 @O_CodigoFormaDePago		Int,
	 @O_Autorizado				Bit

AS 

 Declare @NroOrden as bigint

	Select @NroOrden = N_NumeroOrdenDeCompra+1 From OC_Numeros Where N_CentroDeCostoEmisor = @O_CentroDeCostoEmisor

	Update OC_Numeros 
		Set N_NumeroOrdenDeCompra = N_NumeroOrdenDeCompra + 1 
	Where N_CentroDeCostoEmisor = @O_CentroDeCostoEmisor

	INSERT INTO OC_OrdenesDeCompraCabecera 
		 (O_NumeroOrdenDeCompra,
		  O_Fecha,
		  O_Responsable,
		  O_CodigoProveedor,
		  O_LugarDeEntrega,
		  O_FormaDePagoPactada,
		  O_EmpresaFacturaANombreDe,
	 	  U_Usuario,
		  O_CentroDeCostoEmisor,
		  O_Observaciones,
		  O_FechaEmision,
		  O_CodigoLugarDeEntrega,
		  O_CodigoFormaDePago,
		  O_Autorizado) 
 	VALUES 
		(@NroOrden,
		 @O_Fecha,
		 @O_Responsable,
		 @O_CodigoProveedor,
		 @O_LugarDeEntrega,
		 @O_FormaDePagoPactada,
		 @O_EmpresaFacturaANombreDe,
	 	 @U_Usuario,
		 @O_CentroDeCostoEmisor,
		 @O_Observaciones,
		 @O_FechaEmision,
		 @O_CodigoLugarDeEntrega,
		 @O_CodigoFormaDePago,
		 @O_Autorizado)
--retorna el Nro de orden de compra
	Select @NroOrden O_NumeroOrdenDeCompra
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE SpOCOrdenesDeCompraCabeceraModificar
	 @O_NumeroOrdenDeCompra bigint,
	 @O_Fecha smalldatetime,
	 @O_Responsable varchar(50),
	 @O_CodigoProveedor int,
	 @O_LugarDeEntrega varchar(50),
	 @O_FormaDePagoPactada varchar(50),
	 @O_EmpresaFacturaANombreDe char(3),
	 @U_Usuario varchar(15),
	 @O_CentroDeCostoEmisor as varchar(4),
	 @O_Observaciones varchar(100),
	 @O_FechaEmision DateTime,
	 @O_CodigoLugarDeEntrega Int,
	 @O_CodigoFormaDePago Int,
	 @O_Autorizado Bit

AS 
--borra las lineas luego insertar
	DELETE OC_OrdenesDeCompraRenglones
	WHERE O_NumeroOrdenDeCompra = @O_NumeroOrdenDeCompra And O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor


	UPDATE OC_OrdenesDeCompraCabecera 
	Set	 O_Fecha = @O_Fecha,
		 O_Responsable = @O_Responsable,
		 O_CodigoProveedor = @O_CodigoProveedor,
		 O_LugarDeEntrega = @O_LugarDeEntrega,
		 O_FormaDePagoPactada = @O_FormaDePagoPactada,
		 O_EmpresaFacturaANombreDe = @O_EmpresaFacturaANombreDe,
	 	 U_Usuario = @U_Usuario,		 
		 O_Observaciones = @O_Observaciones,
		 O_FechaEmision = @O_FechaEmision,
		 O_CodigoLugarDeEntrega = @O_CodigoLugarDeEntrega,
		 O_CodigoFormaDePago = @O_CodigoFormaDePago,
		 O_Autorizado = @O_Autorizado
 	Where   O_NumeroOrdenDeCompra = @O_NumeroOrdenDeCompra And
		O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE PROCEDURE SpOCOrdenesDeCompraCabeceraTraer
				 @FechaDesde as DateTime,
				 @FechaHasta as DateTime,
				 @Usuario as nvarchar(25),
	 			 @O_CentroDeCostoEmisor varchar(4)
AS 
Declare @N_Nivel as int
	Select @N_Nivel = N_Nivel from AC_AccesosDeGruposAModulos Where M_Modulo = 'A012200' and
		G_Grupo in (Select G_Grupo from AC_Usuarios Where U_Usuario = @Usuario)

  IF @N_Nivel=2 
    BEGIN	
        IF  @O_CentroDeCostoEmisor =''
	Select OC_Cab.O_NumeroOrdenDeCompra, OC_Cab.O_Fecha, 
      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, OC_Cab.O_CentroDeCostoEmisor,
                        OC_Cab.O_LugarDeEntrega, OC_Cab.O_FormaDePagoPactada, 
                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario,
                        Emp.E_Descripcion, OC_Cab.O_FechaAnulacion, O_Autorizado
             From OC_OrdenesDeCompraCabecera AS OC_Cab 
	LEFT JOIN
                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
	Where O_Fecha Between @FechaDesde And @FechaHasta
	Order By O_Fecha
       ELSE
	Select OC_Cab.O_NumeroOrdenDeCompra, OC_Cab.O_Fecha, 
      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor,  OC_Cab.O_CentroDeCostoEmisor,
                        OC_Cab.O_LugarDeEntrega, OC_Cab.O_FormaDePagoPactada, 
                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario,
                        Emp.E_Descripcion, OC_Cab.O_FechaAnulacion, O_Autorizado
             From OC_OrdenesDeCompraCabecera AS OC_Cab 
	 LEFT JOIN
                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
	Where O_Fecha Between @FechaDesde And @FechaHasta And O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor
	Order By O_Fecha
   END
  ELSE
   BEGIN
	Select OC_Cab.O_NumeroOrdenDeCompra, OC_Cab.O_Fecha, 
      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor,  OC_Cab.O_CentroDeCostoEmisor,
                        OC_Cab.O_LugarDeEntrega, OC_Cab.O_FormaDePagoPactada, 
                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario,
                        Emp.E_Descripcion, OC_Cab.O_FechaAnulacion
             From OC_OrdenesDeCompraCabecera AS OC_Cab 
	 LEFT JOIN
                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
	Where O_Fecha Between @FechaDesde And @FechaHasta And O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor-- and U_Usuario = @Usuario
	Order By O_Fecha
 END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeCompraCabeceraTraerNro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeCompraCabeceraTraerNro]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE SpOCOrdenesDeCompraCabeceraTraerNro	 
				@NroOrden as int,
 				@Usuario as nvarchar(25),
				@O_CentroDeCostoEmisor varchar(4)
AS 

Declare @N_Nivel as int
	Select @N_Nivel = N_Nivel from AC_AccesosDeGruposAModulos Where M_Modulo = 'A012100' and
		G_Grupo in (Select G_Grupo from AC_Usuarios Where U_Usuario = @Usuario)

     IF @N_Nivel=2 
     BEGIN	
	Select OC_Cab.O_NumeroOrdenDeCompra, OC_Cab.O_Fecha, OC_Cab.O_Observaciones,
      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor,  O_CodigoLugarDeEntrega,
                        OC_Cab.O_LugarDeEntrega, OC_Cab.O_FormaDePagoPactada, O_CodigoFormaDePago,
                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario, O_FechaEmision,
                        Emp.E_Descripcion, OC_Cab.O_CentroDeCostoEmisor, OC_Cab.O_FechaAnulacion, O_Autorizado
             From OC_OrdenesDeCompraCabecera AS OC_Cab 
	         LEFT  JOIN
                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
	Where O_NumeroOrdenDeCompra = @NroOrden And O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor
   END
   ELSE
   BEGIN
	Select OC_Cab.O_NumeroOrdenDeCompra, OC_Cab.O_Fecha, OC_Cab.O_Observaciones,
      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, O_CodigoLugarDeEntrega,
                        OC_Cab.O_LugarDeEntrega, OC_Cab.O_FormaDePagoPactada, O_CodigoFormaDePago,
                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario, O_FechaEmision,
                        Emp.E_Descripcion, OC_Cab.O_CentroDeCostoEmisor, OC_Cab.O_FechaAnulacion, O_Autorizado
             From OC_OrdenesDeCompraCabecera AS OC_Cab 
	         LEFT  JOIN
                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
	Where O_NumeroOrdenDeCompra = @NroOrden --And U_Usuario=@Usuario
		 And O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor

   END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeCompraCabeceraAutorizar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeCompraCabeceraAutorizar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeCompraCabeceraAutorizarTraer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeCompraCabeceraAutorizarTraer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeContratacionCabeceraAgregar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeContratacionCabeceraAgregar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeContratacionCabeceraModificar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeContratacionCabeceraModificar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeContratacionCabeceraTraer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeContratacionCabeceraTraer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOCOrdenesDeContratacionCabeceraTraerNro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOCOrdenesDeContratacionCabeceraTraerNro]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOcOrdenesDeContratacionCabeceraAutorizar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOcOrdenesDeContratacionCabeceraAutorizar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOcOrdenesDeContratacionCabeceraAutorizarTraer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOcOrdenesDeContratacionCabeceraAutorizarTraer]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE SpOCOrdenesDeCompraCabeceraAutorizar
	 @O_NumeroOrdenDeCompra bigint,	
	 @O_CentroDeCostoEmisor as varchar(4),
	 @O_UsuarioAutorizo VarChar(15)

AS 
	UPDATE OC_OrdenesDeCompraCabecera 
	Set O_Autorizado = 1,
	      O_UsuarioAutorizo = @O_UsuarioAutorizo,
	      O_FechaAutorizo = GetDate()
 	Where O_NumeroOrdenDeCompra = @O_NumeroOrdenDeCompra And
	            O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE SpOCOrdenesDeCompraCabeceraAutorizarTraer
				 @FechaDesde as DateTime,
				 @FechaHasta as DateTime,
				 @CentroDeCostoEmisor varchar(4)
AS 
	IF @CentroDeCostoEmisor ='' 
		BEGIN
			Select OC_Cab.O_NumeroOrdenDeCompra, OC_Cab.O_Fecha, 
		      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, OC_Cab.O_CentroDeCostoEmisor,
		                        OC_Cab.O_LugarDeEntrega, OC_Cab.O_FormaDePagoPactada, 
		                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario,
		                        Emp.E_Descripcion, OC_Cab.O_FechaAnulacion, O_Autorizado
		             From OC_OrdenesDeCompraCabecera AS OC_Cab 
			LEFT JOIN
		                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
			Where O_Fecha Between @FechaDesde And @FechaHasta and O_Autorizado=0 and O_FechaAnulacion is null
			Order By O_Fecha
		END
	ELSE
		BEGIN
			Select OC_Cab.O_NumeroOrdenDeCompra, OC_Cab.O_Fecha, 
		      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, OC_Cab.O_CentroDeCostoEmisor,
		                        OC_Cab.O_LugarDeEntrega, OC_Cab.O_FormaDePagoPactada, 
		                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario,
		                        Emp.E_Descripcion, OC_Cab.O_FechaAnulacion, O_Autorizado
		             From OC_OrdenesDeCompraCabecera AS OC_Cab 
			LEFT JOIN
		                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
			Where O_Fecha Between @FechaDesde And @FechaHasta and O_Autorizado=0 
			    and O_FechaAnulacion is null And O_CentroDeCostoEmisor= @CentroDeCostoEmisor
			Order By O_Fecha
		END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE SpOCOrdenesDeContratacionCabeceraAgregar
	(@O_Fecha 				smalldatetime,
	 @O_Responsable 			varchar(50),
	 @O_CodigoProveedor 			int,
	 @O_LugarDelServicio			varchar(50),
	 @O_FormaDePagoPactada 		varchar(50),
	 @O_EmpresaFacturaANombreDe 	char(3),
	 @U_Usuario 				varchar(15),
	 @O_CentroDeCostoEmisor 		varchar(4),
	 @O_Observaciones			varchar (100) ,
	 @O_FechaEmision			DateTime,
	 @O_CodigoFormaDePago		Int,
	 @O_Autorizado				Bit)

AS 

 Declare @NroOrden as bigint

	Select @NroOrden = N_NumeroDeOrdenDeContratacion+1 From OC_Numeros Where N_CentroDeCostoEmisor = @O_CentroDeCostoEmisor

	Update OC_Numeros 
		Set N_NumeroDeOrdenDeContratacion = N_NumeroDeOrdenDeContratacion + 1
	Where N_CentroDeCostoEmisor = @O_CentroDeCostoEmisor

	INSERT INTO OC_OrdenesDeContratacionCabecera 
		 ( O_NumeroOrdenDeContratacion,
		 O_Fecha,
		 O_Responsable,
		 O_CodigoProveedor,
		 O_LugarDelServicio,
		 O_FormaDePagoPactada,
		 O_EmpresaFacturaANombreDe,
	 	 U_Usuario,
		 O_CentroDeCostoEmisor,
		 O_Observaciones,
		 O_FechaEmision,
		 O_CodigoFormaDePago,
		 O_Autorizado) 
 	VALUES 
		(@NroOrden,
		 @O_Fecha,
		 @O_Responsable,
		 @O_CodigoProveedor,
		 @O_LugarDelServicio,
		 @O_FormaDePagoPactada,
		 @O_EmpresaFacturaANombreDe,
	 	 @U_Usuario,
		 @O_CentroDeCostoEmisor,
		 @O_Observaciones,
		 @O_FechaEmision,
		 @O_CodigoFormaDePago,
		 @O_Autorizado)
--retorna el Nro de orden de Contratación
	Select @NroOrden O_NumeroOrdenDeContratacion
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE SpOCOrdenesDeContratacionCabeceraModificar
	(@O_NumeroOrdenDeContratacion 	bigint,
	 @O_Fecha 				smalldatetime,
	 @O_Responsable 			varchar(50),
	 @O_CodigoProveedor 			int,
	 @O_LugarDelServicio 			varchar(50),
	 @O_FormaDePagoPactada 		varchar(50),
	 @O_EmpresaFacturaANombreDe 	char(3),
	 @U_Usuario 				varchar(15),
	 @O_CentroDeCostoEmisor  		varchar(4),
	 @O_Observaciones			varchar(100),
	 @O_FechaEmision			DateTime,
	 @O_CodigoFormaDePago		Int,	
	 @O_Autorizado				Bit)

AS 

	--borra las línear de una orden	
	DELETE OC_OrdenesDeContratacionRenglones
		WHERE (O_NumeroOrdenDeContratacion = @O_NumeroOrdenDeContratacion And O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor)

	Update OC_OrdenesDeContratacionCabecera 
	Set  O_Fecha = @O_Fecha,
	       O_Responsable = @O_Responsable,
	       O_CodigoProveedor = @O_CodigoProveedor,
	       O_LugarDelServicio = @O_LugarDelServicio, 
	       O_FormaDePagoPactada = @O_FormaDePagoPactada,
	       O_EmpresaFacturaANombreDe = @O_EmpresaFacturaANombreDe,
	       U_Usuario = @U_Usuario,
	       O_Observaciones = @O_Observaciones,
	       O_FechaEmision = @O_FechaEmision,
	       O_CodigoFormaDePago = @O_CodigoFormaDePago,
	       O_Autorizado = @O_Autorizado
 	Where 	O_NumeroOrdenDeContratacion = @O_NumeroOrdenDeContratacion and
		O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE PROCEDURE SpOCOrdenesDeContratacionCabeceraTraer
				 @FechaDesde as DateTime,
				 @FechaHasta as DateTime,
				 @Usuario as nvarchar(25),
				 @O_CentroDeCostoEmisor varchar(4)
AS 
Declare @N_Nivel as int
	Select @N_Nivel = N_Nivel from AC_AccesosDeGruposAModulos Where M_Modulo = 'A016100' and
		G_Grupo in (Select G_Grupo from AC_Usuarios Where U_Usuario = @Usuario)

     IF @N_Nivel=2 
     BEGIN
         IF @O_CentroDeCostoEmisor =0	
	Select OC_Cab.O_NumeroOrdenDeContratacion, OC_Cab.O_Fecha, 
      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, OC_Cab.O_CentroDeCostoEmisor,
                        OC_Cab.O_LugarDelServicio, OC_Cab.O_FormaDePagoPactada, 
                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario,
                        Emp.E_Descripcion, OC_Cab.O_FechaAnulacion
             From OC_OrdenesDeContratacionCabecera AS OC_Cab 
	LEFT JOIN
                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
	Where O_Fecha Between @FechaDesde And @FechaHasta
	Order By O_NumeroOrdenDeContratacion
        ELSE
	Select OC_Cab.O_NumeroOrdenDeContratacion, OC_Cab.O_Fecha, 
      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, OC_Cab.O_CentroDeCostoEmisor,
                        OC_Cab.O_LugarDelServicio, OC_Cab.O_FormaDePagoPactada, 
                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario,
                         Emp.E_Descripcion, OC_Cab.O_FechaAnulacion
             From OC_OrdenesDeContratacionCabecera AS OC_Cab 
	 LEFT JOIN
                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
	Where O_Fecha Between @FechaDesde And @FechaHasta and OC_Cab.O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor
	Order By O_NumeroOrdenDeContratacion

   END
   ELSE
   BEGIN
	Select OC_Cab.O_NumeroOrdenDeContratacion, OC_Cab.O_Fecha, 
      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, OC_Cab.O_CentroDeCostoEmisor,
                        OC_Cab.O_LugarDelServicio, OC_Cab.O_FormaDePagoPactada, 
                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario,
                        Emp.E_Descripcion, OC_Cab.O_FechaAnulacion
             From OC_OrdenesDeContratacionCabecera AS OC_Cab 
	LEFT JOIN
                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
	Where O_Fecha Between @FechaDesde And @FechaHasta and OC_Cab.O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor--and U_Usuario = @Usuario
	Order By O_NumeroOrdenDeContratacion
 END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE PROCEDURE SpOCOrdenesDeContratacionCabeceraTraerNro	 
				@NroOrden as int,
 				@Usuario as nvarchar(25),
				@O_CentroDeCostoEmisor varchar(4)
AS 

Declare @N_Nivel as int
	Select @N_Nivel = N_Nivel from AC_AccesosDeGruposAModulos Where M_Modulo = 'A016100' and
		G_Grupo in (Select G_Grupo from AC_Usuarios Where U_Usuario = @Usuario)

     IF @N_Nivel=2 
	     BEGIN	
		Select OC_Cab.O_NumeroOrdenDeContratacion, OC_Cab.O_Fecha, 
	      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, OC_Cab.O_Observaciones,
	                        OC_Cab.O_LugarDelServicio, OC_Cab.O_FormaDePagoPactada, O_CodigoFormaDePago,
	                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario, O_FechaEmision,
	                        Emp.E_Descripcion, OC_Cab.O_CentroDeCostoEmisor, OC_Cab.O_FechaAnulacion, O_Autorizado
	             From OC_OrdenesDeContratacionCabecera AS OC_Cab  LEFT JOIN
	                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
		Where O_NumeroOrdenDeContratacion = @NroOrden And O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor
	   END
   ELSE
	   BEGIN
		Select OC_Cab.O_NumeroOrdenDeContratacion, OC_Cab.O_Fecha, 
	      	           OC_Cab.O_Responsable, OC_Cab.O_CodigoProveedor, OC_Cab.O_Observaciones,
	                        OC_Cab.O_LugarDelServicio, OC_Cab.O_FormaDePagoPactada, O_CodigoFormaDePago,
	                        OC_Cab.O_EmpresaFacturaANombreDe, OC_Cab.U_Usuario, O_FechaEmision,
	                        Emp.E_Descripcion, OC_Cab.O_CentroDeCostoEmisor, OC_Cab.O_FechaAnulacion, O_Autorizado
	             From OC_OrdenesDeContratacionCabecera AS OC_Cab LEFT JOIN
	                      TA_Empresas AS Emp On OC_Cab.O_EmpresaFacturaANombreDe = Emp.E_Codigo
		Where O_NumeroOrdenDeContratacion = @NroOrden  And O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor
	
	   END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE SpOcOrdenesDeContratacionCabeceraAutorizar
	 @O_NumeroOrdenDeContratacion bigint,	
	 @O_CentroDeCostoEmisor as varchar(4),
	 @O_UsuarioAutorizo VarChar(15)

AS 
	UPDATE OC_OrdenesDeContratacionCabecera 
	Set O_Autorizado = 1,
	      O_UsuarioAutorizo = @O_UsuarioAutorizo,
	      O_FechaAutorizo = GetDate()
 	Where O_NumeroOrdenDeContratacion = @O_NumeroOrdenDeContratacion And
	            O_CentroDeCostoEmisor = @O_CentroDeCostoEmisor

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE SpOcOrdenesDeContratacionCabeceraAutorizarTraer
				 @FechaDesde as DateTime,
				 @FechaHasta as DateTime				,
				 @CentroDeCostoEmisor varchar(4)
AS 
	IF @CentroDeCostoEmisor =''
		Begin
			Select O_NumeroOrdenDeContratacion, O_Fecha, O_CentroDeCostoEmisor, O_Responsable,
		                        O_CodigoProveedor, O_EmpresaFacturaANombreDe, U_Usuario, O_FechaAnulacion,
		                        O_Observaciones, O_FechaEmision, O_CodigoFormaDePago, O_Autorizado,
			          O_UsuarioAutorizo, O_FechaAutorizo
		             From OC_OrdenesDeContratacionCabecera AS OC_Cab 	
			Where O_Fecha Between @FechaDesde And @FechaHasta and O_Autorizado=0 and O_FechaAnulacion is null
			Order By O_Fecha
		End
	ELSE
		Begin
			Select O_NumeroOrdenDeContratacion, O_Fecha, O_CentroDeCostoEmisor, O_Responsable,
		                        O_CodigoProveedor, O_EmpresaFacturaANombreDe, U_Usuario, O_FechaAnulacion,
		                        O_Observaciones, O_FechaEmision, O_CodigoFormaDePago, O_Autorizado,
			          O_UsuarioAutorizo, O_FechaAutorizo
		             From OC_OrdenesDeContratacionCabecera AS OC_Cab 	
			Where O_Fecha Between @FechaDesde And @FechaHasta and O_Autorizado=0 and 
			            O_FechaAnulacion is null and O_CentroDeCostoEmisor = @CentroDeCostoEmisor
			Order By O_Fecha
		End
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOcOrdenesDeContratacionSinPagarTraer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOcOrdenesDeContratacionSinPagarTraer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpOcOrdenesPendientesDeArticuloTraer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SpOcOrdenesPendientesDeArticuloTraer]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE SpOcOrdenesDeContratacionSinPagarTraer 
					@P_Codigo  int
AS

	SELECT Reng.O_NumeroOrdenDeContratacion, O_CuentaContable, O_CentroDeCosto, O_PrecioPactado, 
	               Reng.O_CentroDeCostoEmisor, O_Fecha, O_Observaciones, O_Autorizado
	FROM OC_OrdenesDeContratacionRenglones Reng
 	Join (Select O_NumeroOrdenDeContratacion,  O_CentroDeCostoEmisor, O_Fecha, O_Observaciones, O_Autorizado
	        From OC_OrdenesDeContratacionCabecera
	        Where O_CodigoProveedor = @P_Codigo and O_FechaAnulacion is null ) Cab
	On Cab.O_CentroDeCostoEmisor = Reng.O_CentroDeCostoEmisor and Cab.O_NumeroOrdenDeContratacion = Reng.O_NumeroOrdenDeContratacion
	Where O_PendienteAutorizacionDePago = 1
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE SpOcOrdenesPendientesDeArticuloTraer  
			@P_Codigo as int,
			@A_Codigo as int
AS
		    Select Reng.O_NumeroOrdenDeCompra, Max(O_FormaDePagoPactada) O_FormaDePagoPactada , Reng.O_CentroDeCostoEmisor, Sum(O_CantidadPedida) O_CantidadPedida, 
			Sum(O_CantidadPendiente) O_CantidadPendiente, Avg(O_PrecioPactado) O_PrecioPactado, Min(O_Fecha) O_Fecha
		    From OC_OrdenesDeCompraRenglones Reng
		    Join (Select O_NumeroOrdenDeCompra, O_CentroDeCostoEmisor, O_Fecha ,O_FormaDePagoPactada
				From OC_OrdenesDeCompraCabecera 
				Where O_CodigoProveedor = @P_Codigo and O_FechaAnulacion Is Null and O_Autorizado=1) Cab
		   ON Cab.O_NumeroOrdenDeCompra = Reng.O_NumeroOrdenDeCompra and Cab.O_CentroDeCostoEmisor = Reng.O_CentroDeCostoEmisor
		    Where O_CodigoArticulo = @A_Codigo And O_CantidadPendiente > 0 
		   Group By Reng.O_NumeroOrdenDeCompra, Reng.O_CentroDeCostoEmisor
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

