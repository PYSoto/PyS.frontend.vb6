/*
SQLyog Ultimate v12.09 (32 bit)
MySQL - 5.7.22-0ubuntu0.16.04.1 : Database - modelopys
*********************************************************************
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;
CREATE DATABASE /*!32312 IF NOT EXISTS*/`modelopys` /*!40100 DEFAULT CHARACTER SET latin1 */;

USE `modelopys`;

/*Table structure for table `articulos` */

DROP TABLE IF EXISTS `articulos`;

CREATE TABLE `articulos` (
  `Codigo` varchar(20) NOT NULL DEFAULT '0',
  `Descripcion` varchar(50) NOT NULL DEFAULT '',
  `Art_PrecioVentaConIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Art_PrecioVentaSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Art_PrecioListaConIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Art_PrecioListaSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Art_PrecioCompraSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Art_PrecioCompraSinIVAAnterior` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Iva105` tinyint(1) NOT NULL DEFAULT '0',
  `Exento` tinyint(1) NOT NULL DEFAULT '0',
  `Art_ModeloCamion` varchar(50) NOT NULL DEFAULT '',
  `Art_FechaActualizacion` date NOT NULL,
  `Art_Origen` varchar(5) NOT NULL DEFAULT '',
  `Art_Descuento` varchar(5) NOT NULL DEFAULT '',
  `Art_Prv_ID` int(4) NOT NULL DEFAULT '0',
  `Art_UltimaCompra` date NOT NULL,
  `Art_Marca` varchar(50) NOT NULL DEFAULT '',
  `Art_Catalogo` varchar(50) NOT NULL DEFAULT '',
  `precioListaSinIvaUSD` decimal(16,2) NOT NULL DEFAULT '0.00',
  `cotizacionID` int(4) NOT NULL DEFAULT '0',
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`Codigo`),
  KEY `Descripcion` (`Descripcion`),
  KEY `Art_Prv_ID` (`Art_Prv_ID`),
  KEY `Clave` (`Clave`),
  KEY `cotizacionID` (`cotizacionID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `articulosalias` */

DROP TABLE IF EXISTS `articulosalias`;

CREATE TABLE `articulosalias` (
  `clave` int(4) NOT NULL AUTO_INCREMENT,
  `AAl_Art_ID` varchar(20) NOT NULL DEFAULT '',
  `AAl_Alias` varchar(20) NOT NULL DEFAULT '',
  `AAl_Prv_ID` int(4) NOT NULL DEFAULT '0',
  `AAl_PrecioCompra` decimal(16,2) NOT NULL DEFAULT '0.00',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`clave`),
  UNIQUE KEY `AAl_Alias` (`AAl_Alias`),
  KEY `AAl_Art_ID` (`AAl_Art_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `articulosalternativo` */

DROP TABLE IF EXISTS `articulosalternativo`;

CREATE TABLE `articulosalternativo` (
  `clave` int(4) NOT NULL AUTO_INCREMENT,
  `ArA_Art_ID` varchar(20) NOT NULL DEFAULT '',
  `ArA_Art_ID_Alternativo` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`clave`),
  UNIQUE KEY `ArA_Art_ID` (`ArA_Art_ID`,`ArA_Art_ID_Alternativo`),
  KEY `ArA_Art_ID_Alternativo` (`ArA_Art_ID_Alternativo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `articulosimport` */

DROP TABLE IF EXISTS `articulosimport`;

CREATE TABLE `articulosimport` (
  `Art_Fecha` date NOT NULL,
  `Art_Codigo` varchar(20) NOT NULL DEFAULT '0',
  `Art_Descripcion` varchar(50) NOT NULL DEFAULT '',
  `Art_PrecioListaSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Art_Origen` varchar(5) NOT NULL DEFAULT '',
  `Art_Descuento` varchar(5) NOT NULL DEFAULT '',
  `Art_FechaActualizacion` date NOT NULL,
  `cotizacion_id` int(4) NOT NULL DEFAULT '0',
  `valor_usd` decimal(6,2) NOT NULL DEFAULT '0.00',
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`Art_Fecha`,`Art_Codigo`),
  KEY `Descripcion` (`Art_Descripcion`),
  KEY `Clave` (`Clave`),
  KEY `cotizacion_id` (`cotizacion_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `articulosmirror` */

DROP TABLE IF EXISTS `articulosmirror`;

CREATE TABLE `articulosmirror` (
  `arm_dsn` varchar(20) NOT NULL DEFAULT '',
  `arm_uid` varchar(30) NOT NULL DEFAULT '',
  `arm_pwd` varchar(30) NOT NULL DEFAULT '',
  `arm_ip` varchar(30) NOT NULL DEFAULT '',
  `arm_database` varchar(30) NOT NULL DEFAULT '',
  `arm_id` smallint(2) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`arm_id`),
  KEY `arm_id` (`arm_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `articulosubic` */

DROP TABLE IF EXISTS `articulosubic`;

CREATE TABLE `articulosubic` (
  `AUb_Art_ID` varchar(20) NOT NULL,
  `AUb_Ubicacion` varchar(150) NOT NULL,
  `AUb_ID` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`AUb_Art_ID`),
  KEY `AUb_ID` (`AUb_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `cgosvalores` */

DROP TABLE IF EXISTS `cgosvalores`;

CREATE TABLE `cgosvalores` (
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `CVa_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `codigo` smallint(2) NOT NULL DEFAULT '0',
  `Concepto` varchar(50) NOT NULL DEFAULT '',
  `Numerable` tinyint(1) NOT NULL DEFAULT '0',
  `Duplicados` tinyint(1) NOT NULL DEFAULT '0',
  `FechaEmi` tinyint(1) NOT NULL DEFAULT '0',
  `FechaVto` tinyint(1) NOT NULL DEFAULT '0',
  `Titular` tinyint(1) NOT NULL DEFAULT '0',
  `Banco` tinyint(1) NOT NULL DEFAULT '0',
  `ChTercero` tinyint(1) NOT NULL DEFAULT '0',
  `CtaCte` tinyint(1) NOT NULL DEFAULT '0',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`Clave`),
  UNIQUE KEY `codigo` (`codigo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `clientes` */

DROP TABLE IF EXISTS `clientes`;

CREATE TABLE `clientes` (
  `clave` int(4) NOT NULL AUTO_INCREMENT,
  `Cli_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `Codigo` int(4) NOT NULL DEFAULT '0',
  `cuit` varchar(15) NOT NULL DEFAULT '',
  `razon` varchar(60) NOT NULL DEFAULT '',
  `domicilio` varchar(100) NOT NULL DEFAULT '',
  `Cli_Localidad` varchar(100) NOT NULL DEFAULT '',
  `Cli_Provincia` varchar(100) NOT NULL DEFAULT '',
  `tel` varchar(50) NOT NULL DEFAULT '',
  `fax` varchar(20) NOT NULL DEFAULT '',
  `email` varchar(40) NOT NULL DEFAULT '',
  `Celular` varchar(40) NOT NULL DEFAULT '',
  `Posicion` smallint(2) NOT NULL DEFAULT '0',
  `TipoDoc` varchar(50) NOT NULL DEFAULT '',
  `NroDoc` int(4) NOT NULL DEFAULT '0',
  `LimiteCredito` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Nacionalidad` varchar(50) NOT NULL DEFAULT '',
  `Cli_Descuento` decimal(16,2) NOT NULL DEFAULT '0.00',
  `cli_facturable` tinyint(1) NOT NULL DEFAULT '1',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`clave`),
  UNIQUE KEY `Codigo` (`Cli_Neg_ID`,`Codigo`),
  KEY `cuit` (`cuit`),
  KEY `Posicion` (`Posicion`),
  KEY `razon` (`razon`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `compafip` */

DROP TABLE IF EXISTS `compafip`;

CREATE TABLE `compafip` (
  `caf_id` smallint(2) NOT NULL DEFAULT '0',
  `caf_nombre` varchar(150) NOT NULL DEFAULT '',
  `caf_label` varchar(150) NOT NULL,
  PRIMARY KEY (`caf_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `consfecha` */

DROP TABLE IF EXISTS `consfecha`;

CREATE TABLE `consfecha` (
  `CFe_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `CFe_Fecha` date NOT NULL,
  `CFe_Cuenta` int(4) NOT NULL DEFAULT '0',
  `CFe_Deudor` decimal(16,2) DEFAULT NULL,
  `CFe_Acreedor` decimal(16,2) DEFAULT NULL,
  PRIMARY KEY (`CFe_Neg_ID`,`CFe_Fecha`,`CFe_Cuenta`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `cotizacion` */

DROP TABLE IF EXISTS `cotizacion`;

CREATE TABLE `cotizacion` (
  `cotizacion_id` int(4) NOT NULL AUTO_INCREMENT,
  `fecha` date NOT NULL,
  `usd_compra` decimal(6,2) NOT NULL DEFAULT '0.00',
  `usd_venta` decimal(6,2) NOT NULL DEFAULT '0.00',
  PRIMARY KEY (`cotizacion_id`),
  UNIQUE KEY `fecha` (`fecha`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `detartic` */

DROP TABLE IF EXISTS `detartic`;

CREATE TABLE `detartic` (
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `ClaveMovClie` int(4) NOT NULL DEFAULT '0',
  `CgoComprob` smallint(2) NOT NULL DEFAULT '0',
  `Item` smallint(2) NOT NULL DEFAULT '0',
  `DeA_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `CgoArtic` varchar(20) NOT NULL DEFAULT '',
  `DeA_Descripcion` varchar(150) NOT NULL DEFAULT '',
  `Cantidad` decimal(10,3) NOT NULL DEFAULT '0.000',
  `PrecioVentaSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `PrecioVentaConIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `PrecioDescuentoSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `PrecioDescuentoConIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Iva105` tinyint(1) NOT NULL DEFAULT '0',
  `Exento` tinyint(1) NOT NULL DEFAULT '0',
  `Fecha` date DEFAULT NULL,
  `fechafac` date DEFAULT NULL,
  `PrecioCompraSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `DeA_Descuento` decimal(16,2) NOT NULL DEFAULT '0.00',
  PRIMARY KEY (`Clave`),
  KEY `ClaveMovClie` (`ClaveMovClie`),
  KEY `CgoArtic` (`CgoArtic`(10)),
  KEY `Fecha` (`Fecha`),
  KEY `CgoComprob` (`CgoComprob`),
  KEY `ClaveFIFO` (`CgoArtic`,`Fecha`,`Clave`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `detpresupuesto` */

DROP TABLE IF EXISTS `detpresupuesto`;

CREATE TABLE `detpresupuesto` (
  `DeP_Pre_ID` int(4) NOT NULL DEFAULT '0',
  `DeP_Orden` smallint(2) NOT NULL DEFAULT '0',
  `DeP_Art_ID` varchar(20) NOT NULL DEFAULT '',
  `DeP_Cant_Art` smallint(2) NOT NULL DEFAULT '0',
  `DeP_UnitSIva` decimal(16,2) NOT NULL,
  `DeP_UnitCIva` decimal(16,2) NOT NULL,
  `clave` int(4) NOT NULL AUTO_INCREMENT,
  PRIMARY KEY (`DeP_Pre_ID`,`DeP_Orden`),
  KEY `clave` (`clave`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `empresa` */

DROP TABLE IF EXISTS `empresa`;

CREATE TABLE `empresa` (
  `Emp_ID` smallint(2) NOT NULL AUTO_INCREMENT,
  `Emp_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `NOMBRE` varchar(50) NOT NULL DEFAULT '',
  `Emp_RSocial` varchar(100) NOT NULL DEFAULT '',
  `DOMICILIO` varchar(100) NOT NULL DEFAULT '',
  `TELF` varchar(50) NOT NULL DEFAULT '',
  `CUIT` varchar(13) NOT NULL DEFAULT '',
  `PuntoVta` smallint(2) NOT NULL DEFAULT '0',
  `IngBrutos` varchar(50) NOT NULL DEFAULT '',
  `NroEstablecimiento` varchar(50) NOT NULL DEFAULT '',
  `SedeTimbrado` varchar(50) NOT NULL DEFAULT '',
  `InicioActividades` varchar(30) NOT NULL DEFAULT '',
  `CondicionIva` varchar(30) NOT NULL DEFAULT '',
  `emp_ubicacion` varchar(50) NOT NULL DEFAULT '',
  PRIMARY KEY (`Emp_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `factcobradas` */

DROP TABLE IF EXISTS `factcobradas`;

CREATE TABLE `factcobradas` (
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `ClaveMovC` int(4) NOT NULL DEFAULT '0',
  `ClavePago` int(4) NOT NULL DEFAULT '0',
  `Importe` decimal(16,2) NOT NULL DEFAULT '0.00',
  PRIMARY KEY (`Clave`),
  KEY `ClavePago` (`ClavePago`),
  KEY `ClaveMovp` (`ClaveMovC`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `feriados` */

DROP TABLE IF EXISTS `feriados`;

CREATE TABLE `feriados` (
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `Fecha` date DEFAULT NULL,
  `Concepto` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`Clave`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `importaciones` */

DROP TABLE IF EXISTS `importaciones`;

CREATE TABLE `importaciones` (
  `imp_id` int(4) NOT NULL AUTO_INCREMENT,
  `imp_fecha` date NOT NULL,
  PRIMARY KEY (`imp_id`),
  KEY `imp_fecha` (`imp_fecha`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `movclie` */

DROP TABLE IF EXISTS `movclie`;

CREATE TABLE `movclie` (
  `clave` int(4) NOT NULL AUTO_INCREMENT,
  `MCl_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `MCl_Emp_ID` smallint(2) NOT NULL DEFAULT '1',
  `CgoClie` int(4) NOT NULL DEFAULT '0',
  `CgoComprob` smallint(2) NOT NULL DEFAULT '0',
  `FechaComprob` date NOT NULL,
  `MCl_FechaVenc` date NOT NULL,
  `Prefijo` smallint(2) NOT NULL DEFAULT '0',
  `NroComprob` int(4) NOT NULL DEFAULT '0',
  `Importe` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Cancelado` decimal(16,2) NOT NULL DEFAULT '0.00',
  `NetoSinDescuento` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Neto` decimal(16,2) NOT NULL DEFAULT '0.00',
  `NetoCancelado` decimal(16,2) NOT NULL DEFAULT '0.00',
  `MontoIva` decimal(16,2) NOT NULL DEFAULT '0.00',
  `MontoExento` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Recibo` tinyint(1) NOT NULL DEFAULT '0',
  `Anulada` tinyint(1) NOT NULL DEFAULT '0',
  `TipoCompro` char(1) NOT NULL,
  `MCl_Letras` varchar(250) NOT NULL DEFAULT '',
  `MCl_IVA` decimal(10,2) NOT NULL DEFAULT '21.00',
  `MCl_Observaciones` text NOT NULL,
  `MCl_CAE` varchar(50) NOT NULL DEFAULT '',
  `MCl_CAEVenc` varchar(20) NOT NULL DEFAULT '',
  `MCl_Barras` varchar(40) NOT NULL DEFAULT '',
  PRIMARY KEY (`clave`),
  KEY `CgoComprob` (`CgoComprob`),
  KEY `CgoClie` (`CgoClie`),
  KEY `comprobunico` (`CgoComprob`,`Prefijo`,`NroComprob`),
  KEY `anulada` (`Anulada`),
  KEY `fechacomprob2` (`FechaComprob`,`Prefijo`,`NroComprob`),
  KEY `recibo` (`Recibo`),
  KEY `tipocompro` (`TipoCompro`,`Prefijo`,`NroComprob`),
  KEY `FechaComprob` (`FechaComprob`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `movprov` */

DROP TABLE IF EXISTS `movprov`;

CREATE TABLE `movprov` (
  `CgoProv` int(4) NOT NULL DEFAULT '0',
  `CgoComprob` smallint(2) NOT NULL DEFAULT '0',
  `Prefijo` smallint(2) NOT NULL DEFAULT '0',
  `NroComprob` int(4) NOT NULL DEFAULT '0',
  `MPr_Emp_ID` smallint(2) NOT NULL DEFAULT '1',
  `MPr_Neg_ID` smallint(2) NOT NULL DEFAULT '0',
  `FechaComprob` date NOT NULL,
  `Importe` decimal(10,2) NOT NULL DEFAULT '0.00',
  `Neto` decimal(10,2) NOT NULL DEFAULT '0.00',
  `MontoIva` decimal(10,2) NOT NULL DEFAULT '0.00',
  `MontoIva27` decimal(10,2) NOT NULL DEFAULT '0.00',
  `MontoIva105` decimal(10,2) NOT NULL DEFAULT '0.00',
  `PercIva` decimal(10,2) NOT NULL DEFAULT '0.00',
  `PercIngBrutos` decimal(10,2) NOT NULL DEFAULT '0.00',
  `GNG` decimal(10,2) NOT NULL DEFAULT '0.00',
  `clave` int(4) NOT NULL AUTO_INCREMENT,
  PRIMARY KEY (`CgoProv`,`CgoComprob`,`Prefijo`,`NroComprob`),
  KEY `CgoComprob` (`CgoComprob`),
  KEY `clave` (`clave`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `negocio` */

DROP TABLE IF EXISTS `negocio`;

CREATE TABLE `negocio` (
  `Neg_ID` smallint(2) NOT NULL AUTO_INCREMENT,
  `Neg_Nombre` varchar(200) NOT NULL DEFAULT '',
  `Neg_Domicilio` varchar(250) NOT NULL DEFAULT '',
  `Neg_DSN` varchar(50) NOT NULL DEFAULT '',
  `Neg_IP` varchar(50) NOT NULL DEFAULT '',
  `Neg_DB` varchar(50) NOT NULL DEFAULT '',
  `Neg_User` varchar(50) NOT NULL DEFAULT '',
  PRIMARY KEY (`Neg_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `parametros` */

DROP TABLE IF EXISTS `parametros`;

CREATE TABLE `parametros` (
  `clave` int(4) NOT NULL AUTO_INCREMENT,
  `Par_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `Iva1` decimal(16,2) NOT NULL,
  `Iva2` decimal(16,2) NOT NULL,
  `Par_FEProduccion` tinyint(1) NOT NULL DEFAULT '0',
  `Par_TA` text NOT NULL,
  PRIMARY KEY (`clave`),
  KEY `clave` (`clave`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `presupuesto` */

DROP TABLE IF EXISTS `presupuesto`;

CREATE TABLE `presupuesto` (
  `Pre_ID` int(4) NOT NULL AUTO_INCREMENT,
  `Pre_Cli_ID` int(4) NOT NULL DEFAULT '0',
  `Pre_Fecha` date NOT NULL,
  `Pre_FechaVto` date NOT NULL,
  `Pre_Observac` varchar(250) NOT NULL DEFAULT '',
  `Pre_ClaveMovClie` int(4) NOT NULL,
  `Pre_CtaCte` tinyint(1) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Pre_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `proveedor` */

DROP TABLE IF EXISTS `proveedor`;

CREATE TABLE `proveedor` (
  `proveedor_id` int(4) NOT NULL DEFAULT '0',
  `razon_social` varchar(100) NOT NULL DEFAULT '',
  `nombre_fantasia` varchar(100) NOT NULL DEFAULT '',
  `cuit` varchar(20) NOT NULL DEFAULT '',
  `domicilio` varchar(100) NOT NULL DEFAULT '',
  `localidad` varchar(100) NOT NULL DEFAULT '',
  `provincia` varchar(100) NOT NULL DEFAULT '',
  `telefono` varchar(100) NOT NULL DEFAULT '',
  `fax` varchar(100) NOT NULL DEFAULT '',
  `email` varchar(100) NOT NULL DEFAULT '',
  `posicion_iva` smallint(2) NOT NULL DEFAULT '0',
  `celular` varchar(100) NOT NULL DEFAULT '',
  `ingresos_brutos` varchar(100) NOT NULL DEFAULT '',
  `contacto` varchar(250) NOT NULL DEFAULT '',
  `observaciones` text NOT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`proveedor_id`),
  KEY `cuit` (`cuit`),
  KEY `razon_social` (`razon_social`),
  KEY `posicion_iva` (`posicion_iva`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `proveedor_movimiento` */

DROP TABLE IF EXISTS `proveedor_movimiento`;

CREATE TABLE `proveedor_movimiento` (
  `proveedor_id` int(4) NOT NULL DEFAULT '0',
  `comprobante_id` smallint(2) NOT NULL DEFAULT '0',
  `prefijo` smallint(2) NOT NULL DEFAULT '0',
  `nro_comprobante` int(4) NOT NULL DEFAULT '0',
  `empresa_id` smallint(2) NOT NULL DEFAULT '0',
  `negocio_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha_comprobante` date NOT NULL,
  `fecha_vencimiento` date DEFAULT NULL,
  `total` decimal(16,2) NOT NULL DEFAULT '0.00',
  `total_cancelado` decimal(16,2) NOT NULL DEFAULT '0.00',
  `neto` decimal(16,2) NOT NULL DEFAULT '0.00',
  `importe_iva1` decimal(16,2) NOT NULL DEFAULT '0.00',
  `importe_iva2` decimal(16,2) NOT NULL DEFAULT '0.00',
  `importe_iva3` decimal(16,2) NOT NULL DEFAULT '0.00',
  `percepcion_iva` decimal(16,2) NOT NULL DEFAULT '0.00',
  `percepcion_iibb` decimal(16,2) NOT NULL DEFAULT '0.00',
  `gastos_no_gravados` decimal(16,2) NOT NULL DEFAULT '0.00',
  `ajustes` decimal(16,2) NOT NULL DEFAULT '0.00',
  `monotributo` tinyint(1) NOT NULL DEFAULT '0',
  `cuenta_movimiento_id` bigint(20) NOT NULL DEFAULT '0',
  `observaciones` text NOT NULL,
  `proveedor_movimiento_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`proveedor_id`,`comprobante_id`,`prefijo`,`nro_comprobante`),
  KEY `proveedor_id` (`proveedor_id`,`comprobante_id`,`prefijo`,`nro_comprobante`),
  KEY `fecha_vencimiento` (`fecha_vencimiento`),
  KEY `proveedor_id_2` (`proveedor_id`,`comprobante_id`,`fecha_comprobante`,`prefijo`,`nro_comprobante`,`total`),
  KEY `comprobante_id` (`comprobante_id`),
  KEY `cuenta_movimiento_id` (`cuenta_movimiento_id`),
  KEY `proveedor_movimiento_id` (`proveedor_movimiento_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `proveedor_pago` */

DROP TABLE IF EXISTS `proveedor_pago`;

CREATE TABLE `proveedor_pago` (
  `proveedor_movimiento_id_deuda` int(4) NOT NULL DEFAULT '0',
  `proveedor_movimiento_id_aplicado` int(4) NOT NULL DEFAULT '0',
  `importe_aplicado` decimal(16,2) NOT NULL DEFAULT '0.00',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`proveedor_movimiento_id_deuda`,`proveedor_movimiento_id_aplicado`),
  KEY `proveedor_movimiento_id_deuda` (`proveedor_movimiento_id_deuda`),
  KEY `proveedor_movimiento_id_aplicado` (`proveedor_movimiento_id_aplicado`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `proveedor_saldo` */

DROP TABLE IF EXISTS `proveedor_saldo`;

CREATE TABLE `proveedor_saldo` (
  `proveedor_id` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `saldo` decimal(16,2) NOT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`proveedor_id`,`fecha`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `registrocae` */

DROP TABLE IF EXISTS `registrocae`;

CREATE TABLE `registrocae` (
  `rec_tco_id` smallint(2) NOT NULL DEFAULT '0',
  `rec_prefijo` smallint(2) NOT NULL DEFAULT '0',
  `rec_nrocomprob` int(4) NOT NULL DEFAULT '0',
  `rec_cli_id` int(4) NOT NULL DEFAULT '0',
  `rec_total` decimal(16,2) NOT NULL DEFAULT '0.00',
  `rec_exento` decimal(16,2) NOT NULL DEFAULT '0.00',
  `rec_neto` decimal(16,2) NOT NULL DEFAULT '0.00',
  `rec_neto105` decimal(16,2) NOT NULL DEFAULT '0.00',
  `rec_iva` decimal(16,2) NOT NULL DEFAULT '0.00',
  `rec_iva105` decimal(16,2) NOT NULL DEFAULT '0.00',
  `rec_cae` varchar(30) NOT NULL DEFAULT '',
  `rec_fecha` varchar(20) NOT NULL DEFAULT '',
  `rec_caevenc` varchar(20) NOT NULL DEFAULT '',
  `rec_barras` varchar(50) NOT NULL DEFAULT '',
  `rec_id` int(4) NOT NULL AUTO_INCREMENT,
  PRIMARY KEY (`rec_tco_id`,`rec_prefijo`,`rec_nrocomprob`),
  KEY `rec_id` (`rec_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `tiposcomprob` */

DROP TABLE IF EXISTS `tiposcomprob`;

CREATE TABLE `tiposcomprob` (
  `TCo_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `codigo` int(4) NOT NULL DEFAULT '0',
  `Descripcion` varchar(200) NOT NULL DEFAULT '',
  `Modulo` smallint(2) NOT NULL DEFAULT '0',
  `CtaCte` tinyint(1) NOT NULL DEFAULT '0',
  `Debita` tinyint(1) NOT NULL DEFAULT '0',
  `Iva` tinyint(1) NOT NULL DEFAULT '0',
  `AplicaPend` tinyint(1) NOT NULL DEFAULT '0',
  `Aplicable` tinyint(1) NOT NULL DEFAULT '0',
  `aplicacion` tinyint(1) NOT NULL DEFAULT '0',
  `LibroIva` tinyint(1) NOT NULL DEFAULT '0',
  `TipoComprob` char(1) NOT NULL,
  `Recibo` tinyint(1) NOT NULL DEFAULT '0',
  `Contado` tinyint(1) NOT NULL DEFAULT '0',
  `orden_pago` tinyint(1) NOT NULL DEFAULT '0',
  `TCo_PuntoVta` smallint(2) NOT NULL DEFAULT '0',
  `TCo_TipoAfip` smallint(2) NOT NULL DEFAULT '0',
  `TCo_FactElect` tinyint(1) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`TCo_Neg_ID`,`codigo`),
  UNIQUE KEY `codigo` (`codigo`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `tmbalance` */

DROP TABLE IF EXISTS `tmbalance`;

CREATE TABLE `tmbalance` (
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `Bal_hWnd` int(4) NOT NULL DEFAULT '0',
  `Bal_Cuenta` int(4) NOT NULL DEFAULT '0',
  `Bal_Deudor` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Bal_Acreedor` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Bal_SaldoDeudor` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Bal_SaldoAcreedor` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Bal_Grado` smallint(2) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Clave`),
  KEY `Clave` (`Clave`),
  KEY `Cuenta` (`Bal_Cuenta`),
  KEY `Bal_hWnd` (`Bal_hWnd`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `tmdetartic` */

DROP TABLE IF EXISTS `tmdetartic`;

CREATE TABLE `tmdetartic` (
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `DAT_hWnd` int(4) NOT NULL DEFAULT '0',
  `Item` int(4) NOT NULL DEFAULT '0',
  `CgoArticulo` varchar(20) NOT NULL,
  `Descripcion` varchar(100) NOT NULL,
  `Cantidad` decimal(10,3) NOT NULL DEFAULT '0.000',
  `TotalConIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `TotalSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `PrecioVentaConIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `PrecioVentaSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `PrecioDescuentoConIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `PrecioDescuentoSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `PrecioCompraSinIVA` decimal(16,2) NOT NULL DEFAULT '0.00',
  `Item2` smallint(2) NOT NULL DEFAULT '0',
  `iva105` tinyint(1) NOT NULL DEFAULT '0',
  `exento` tinyint(1) NOT NULL DEFAULT '0',
  `nrofactura` int(4) NOT NULL DEFAULT '0',
  `fechafactura` date DEFAULT NULL,
  `descuento` decimal(16,2) NOT NULL DEFAULT '0.00',
  PRIMARY KEY (`Clave`),
  KEY `Clave` (`Clave`),
  KEY `DAT_hWnd` (`DAT_hWnd`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `usuarioactivo` */

DROP TABLE IF EXISTS `usuarioactivo`;

CREATE TABLE `usuarioactivo` (
  `UAc_IP` char(20) NOT NULL DEFAULT '',
  `UAc_hWnd` int(4) NOT NULL DEFAULT '0',
  `UAc_Login` char(20) DEFAULT NULL,
  `UAc_TimeStamp` timestamp NULL DEFAULT NULL,
  PRIMARY KEY (`UAc_IP`,`UAc_hWnd`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `usuarios` */

DROP TABLE IF EXISTS `usuarios`;

CREATE TABLE `usuarios` (
  `clave` int(4) NOT NULL AUTO_INCREMENT,
  `Nombre` varchar(20) NOT NULL DEFAULT '',
  `Usu_Password` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`clave`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `valores` */

DROP TABLE IF EXISTS `valores`;

CREATE TABLE `valores` (
  `Val_Neg_ID` smallint(2) NOT NULL DEFAULT '1',
  `Codigo` smallint(2) NOT NULL DEFAULT '0',
  `CgoCli` int(4) NOT NULL DEFAULT '0',
  `proveedor_id` int(4) NOT NULL DEFAULT '0',
  `FechaEmi` date DEFAULT NULL,
  `FechaVto` date DEFAULT NULL,
  `Val_TCo_ID` smallint(2) NOT NULL DEFAULT '0',
  `NroComprob` int(4) NOT NULL DEFAULT '0',
  `Importe` decimal(16,2) NOT NULL DEFAULT '0.00',
  `FechaReg` date NOT NULL,
  `ClaveMovV` int(4) NOT NULL DEFAULT '0',
  `proveedormovimiento_id` int(4) NOT NULL DEFAULT '0',
  `Titular` varchar(50) NOT NULL DEFAULT '',
  `Banco` varchar(50) NOT NULL DEFAULT '',
  `Clave` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`Clave`),
  KEY `Codigo` (`Codigo`),
  KEY `CgoCli` (`CgoCli`),
  KEY `Clave` (`Clave`),
  KEY `ClaveMovV` (`ClaveMovV`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `vwaliasprov` */

DROP TABLE IF EXISTS `vwaliasprov`;

/*!50001 DROP VIEW IF EXISTS `vwaliasprov` */;
/*!50001 DROP TABLE IF EXISTS `vwaliasprov` */;

/*!50001 CREATE TABLE  `vwaliasprov`(
 `clave` int(4) ,
 `AAl_Art_ID` varchar(20) ,
 `aal_alias` varchar(20) ,
 `razon` varchar(100) ,
 `aal_preciocompra` decimal(16,2) 
)*/;

/*Table structure for table `vwalterartprov` */

DROP TABLE IF EXISTS `vwalterartprov`;

/*!50001 DROP VIEW IF EXISTS `vwalterartprov` */;
/*!50001 DROP TABLE IF EXISTS `vwalterartprov` */;

/*!50001 CREATE TABLE  `vwalterartprov`(
 `clave` int(4) ,
 `codigo` varchar(20) ,
 `descripcion` varchar(50) ,
 `razon` varchar(100) 
)*/;

/*Table structure for table `vwpresup` */

DROP TABLE IF EXISTS `vwpresup`;

/*!50001 DROP VIEW IF EXISTS `vwpresup` */;
/*!50001 DROP TABLE IF EXISTS `vwpresup` */;

/*!50001 CREATE TABLE  `vwpresup`(
 `pre_id` int(4) ,
 `pre_fecha` date ,
 `pre_cli_id` int(4) ,
 `pre_ctacte` tinyint(1) ,
 `razon` varchar(60) ,
 `NOMBRE` varchar(50) ,
 `TotalSIva` decimal(43,2) ,
 `TotalCIva` decimal(43,2) 
)*/;

/*View structure for view vwaliasprov */

/*!50001 DROP TABLE IF EXISTS `vwaliasprov` */;
/*!50001 DROP VIEW IF EXISTS `vwaliasprov` */;

/*!50001 CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `vwaliasprov` AS (select `articulosalias`.`clave` AS `clave`,`articulosalias`.`AAl_Art_ID` AS `AAl_Art_ID`,`articulosalias`.`AAl_Alias` AS `aal_alias`,`proveedor`.`razon_social` AS `razon`,`articulosalias`.`AAl_PrecioCompra` AS `aal_preciocompra` from (`articulosalias` join `proveedor` on((`articulosalias`.`AAl_Prv_ID` = `proveedor`.`proveedor_id`)))) */;

/*View structure for view vwalterartprov */

/*!50001 DROP TABLE IF EXISTS `vwalterartprov` */;
/*!50001 DROP VIEW IF EXISTS `vwalterartprov` */;

/*!50001 CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `vwalterartprov` AS (select `articulosalternativo`.`clave` AS `clave`,`articulos`.`Codigo` AS `codigo`,`articulos`.`Descripcion` AS `descripcion`,`proveedor`.`razon_social` AS `razon` from ((`articulosalternativo` join `articulos` on((`articulosalternativo`.`ArA_Art_ID_Alternativo` = `articulos`.`Codigo`))) join `proveedor` on((`articulos`.`Art_Prv_ID` = `proveedor`.`proveedor_id`))) order by `articulosalternativo`.`ArA_Art_ID`,`articulos`.`Descripcion`) */;

/*View structure for view vwpresup */

/*!50001 DROP TABLE IF EXISTS `vwpresup` */;
/*!50001 DROP VIEW IF EXISTS `vwpresup` */;

/*!50001 CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `vwpresup` AS select `presupuesto`.`Pre_ID` AS `pre_id`,`presupuesto`.`Pre_Fecha` AS `pre_fecha`,`presupuesto`.`Pre_Cli_ID` AS `pre_cli_id`,`presupuesto`.`Pre_CtaCte` AS `pre_ctacte`,`clientes`.`razon` AS `razon`,`empresa`.`NOMBRE` AS `NOMBRE`,sum((`detpresupuesto`.`DeP_Cant_Art` * `detpresupuesto`.`DeP_UnitSIva`)) AS `TotalSIva`,sum((`detpresupuesto`.`DeP_Cant_Art` * `detpresupuesto`.`DeP_UnitCIva`)) AS `TotalCIva` from (((`presupuesto` join `clientes` on((`clientes`.`Codigo` = `presupuesto`.`Pre_Cli_ID`))) join `detpresupuesto` on((`detpresupuesto`.`DeP_Pre_ID` = `presupuesto`.`Pre_ID`))) join `empresa` on((`empresa`.`Emp_Neg_ID` = `clientes`.`Cli_Neg_ID`))) group by `presupuesto`.`Pre_ID` */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
