# Programa el rastreador en el Programador de tareas de Windows:
# todos los días, cada hora, de 10:00 a 20:00.
# Si mueves esta carpeta, vuelve a ejecutar este archivo para actualizar la ruta.

$nombre = "TechMarleys Seguimiento"
$pythonw = (Get-Command pythonw -ErrorAction Stop).Source
$script = Join-Path $PSScriptRoot "rastreador.py"

$accion = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$script`" --programado" -WorkingDirectory $PSScriptRoot
$disparador = New-ScheduledTaskTrigger -Daily -At "10:00"
$repeticion = New-ScheduledTaskTrigger -Once -At "10:00" -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Hours 10 -Minutes 5)
$disparador.Repetition = $repeticion.Repetition
$opciones = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

Register-ScheduledTask -TaskName $nombre -Action $accion -Trigger $disparador -Settings $opciones `
    -Description "Captura el catálogo de techmarleys.com cada hora (10am-8pm) y actualiza el reporte Excel." -Force | Out-Null

Write-Host "Tarea '$nombre' programada: cada hora de 10:00 a 20:00 usando $script"
