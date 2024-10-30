# Módulo de Corte de Caja Automático en Excel para CANELLA S.A.

Este repositorio contiene un módulo en Visual Basic desarrollado para uso personal para automatizar el proceso de corte de caja en el archivo específico de Excel utilizado por CANELLA S.A. Este módulo procesa automáticamente datos financieros exportados desde SAP en formato Excel (.xlsx) y consolida la información en un formato de reporte prediseñado.



## Disclaimer

Este repositorio **no comparte ninguna información confidencial de CANELLA S.A.**. El código aquí expuesto está diseñado para propósitos demostrativos y **solo funciona en combinación con archivos de reporte de SAP y plantillas Excel específicas de CANELLA S.A.**, que no están incluidos en este repositorio.

## Descripción

El módulo realiza la verificación y consolidación de datos financieros esenciales, incluyendo pagos en efectivo, cheques, tarjetas y depósitos. Además, detecta discrepancias en los totales de los métodos de pago y notifica al cajero sobre cobros en dólares u otras posibles diferencias.

## Requerimientos

Para ejecutar este módulo correctamente, es necesario:
- El archivo Excel de corte de Caja con el modulo incluido (no adjunto: archivo confidencial)
- Contar con el archivo de reporte de SAP exportado en formato `.xlsx`, ubicado en la misma carpeta que el archivo Excel de corte de caja.
- Rellenar manualmente ciertos datos financieros en el reporte, específicamente:
  - Cheques
  - Tarjetas
  - Depósitos
  - Recibos de abono a facturas anuladas


## Funciones Principales

### `CorteDeCaja`
Se inicia el proceso de corte de caja presionando el botón que ejecuta la macro, solicitando la confirmación del usuario antes de proceder y verificando la correcta ubicación del reporte de SAP sin cambios de nombre.


### Cálculo de Totales
Estas funciones calculan los totales para diferentes métodos de pago y categorías:
- **`totalEfectivo`**: Calcula el total de efectivo en facturas y recibos.
- **`totalCheques`**: Calcula el total de cobros en cheques y verifica posibles discrepancias.
- **`totalTarjetas`**: Calcula el total de cobros en tarjetas.
- **`totalDepositos`**: Calcula el total de cobros en depósitos.
- **`totalExcenciones`**: Calcula el total de exenciones.

### Notificaciones de Discrepancias
Genera notificaciones para el cajero si existen diferencias en:
- Cheques (`notificarDiferenciaCheques`)
- Tarjetas (`notificarDiferenciaTarjetas`)
- Depósitos (`notificarDiferenciaDepositos`)

### Notificaciones de Cobros Especiales
- **`cobroDolares`**: Verifica si hubo cobros en dólares y notifica al cajero para que ajuste el tipo de cambio.
- **`notificarCobroDolaresEnRecibosQuetzales`**: Notifica si un recibo de caja en Quetzales fue cobrado en dólares, lo cual requiere corrección.

## Ejemplo de Uso

1. Se debe tener el archivo de reporte de SAP exportado como `.xlsx` en la misma ubicación que el archivo Excel que contiene este módulo.
2. Completa manualmente los datos de cheques, tarjetas, depósitos y recibos en el reporte.
3. Ejecuta el subproceso `CorteDeCaja`por medio de el botón que ejecuta la macro.
4. Revisa las notificaciones de diferencias para verificar la precisión del reporte antes de su envío.


## Consideraciones

El propósito de este módulo es agilizar y facilitar el corte de caja, automatizando el proceso de copiado de datos financieros y reduciendo errores manuales. Sin embargo, es crucial realizar una revisión manual final para confirmar que todos los datos son precisos y correctos.
