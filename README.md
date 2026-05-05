# PrinterAlert - Monitor de Impresión Kanban DOOR

Monitor automático que detecta desincronización entre las secuencias de etiquetas recibidas (archivos `H_`) y las secuencias efectivamente impresas (log de la consola de impresión).

## Problema

La consola de impresión (`print-services-2.4.2-console`) se cuelga ocasionalmente, dejando etiquetas sin imprimir. Este script detecta esa situación y guía/ejecuta las acciones correctivas.

## Uso

```powershell
# Modo demo (solo muestra qué haría)
.\Monitor-PrintKanban.ps1

# Modo interactivo (pregunta antes de cada paso)
.\Monitor-PrintKanban.ps1 -ModoEjecucion "preguntar"

# Modo automático (ejecuta los pasos correctivos)
.\Monitor-PrintKanban.ps1 -ModoEjecucion "auto"

# Cambiar tolerancia (default: 5)
.\Monitor-PrintKanban.ps1 -Tolerancia 3
```

## Parámetros

| Parámetro | Default | Descripción |
|-----------|---------|-------------|
| `-Tolerancia` | `5` | Diferencia máxima permitida entre secuencias |
| `-ModoEjecucion` | `"demo"` | `demo`, `preguntar` o `auto` |
| `-Servidor` | `172.17.132.153` | IP del servidor |
| `-RutaLogs` | `\\172.17.132.153\c$\Program Files\JITMS\prod\print-services-2.4.2-console\logs` | Ruta a los logs |
| `-RutaProcesados` | `\\172.17.132.153\c$\Users\rpareja1ta\Desktop\SFTP\SFTP_TASA\Procesados` | Ruta a archivos H_ |
| `-NombreServicio` | `ImportTXTTASAsvcHost` | Servicio Windows a controlar |
| `-PosicionSecuencia` | `10` | Posición (0-based) de la secuencia en el archivo H_ |
| `-LargoSecuencia` | `3` | Largo del campo de secuencia |

## Acciones correctivas (cuando diferencia > tolerancia)

1. **Detener** servicio `ImportTXTTASAsvcHost`
2. **Reiniciar** la consola de impresión manualmente
3. **Esperar** a que se imprima la última secuencia
4. **Iniciar** servicio `ImportTXTTASAsvcHost`
