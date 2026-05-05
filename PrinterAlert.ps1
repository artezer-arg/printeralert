Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ============================================================
# CONFIGURACION
# ============================================================
$script:Config = @{
    Servidor          = "172.17.132.153"
    RutaLogs          = "\\172.17.132.153\c$\Program Files\JITMS\prod\print-services-2.4.2-console\logs"
    RutaProcesados    = "\\172.17.132.153\c$\Users\rpareja1ta\Desktop\SFTP\SFTP_TASA\Procesados"
    NombreServicio    = "ImportTXTTASAsvcHost"
    PosicionSecuencia = 10
    LargoSecuencia    = 3
    Tolerancia        = 5
    ConsolaExe        = "JIT.Services.PrintMgrSvcHostConsole.exe"
    ConsolaRuta       = "C:\Program Files\JITMS\prod\print-services-2.4.2-console\JIT.Services.PrintMgrSvcHostConsole.exe"
}

# ============================================================
# FUNCIONES DE NEGOCIO
# ============================================================
function Get-SecuenciaRecibida {
    try {
        $ultimo = Get-ChildItem -Path $script:Config.RutaProcesados -Filter "H_*.TXT" -ErrorAction Stop |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $ultimo) { return @{ Secuencia = -1; Archivo = "N/A"; Fecha = "N/A"; Error = "No se encontraron archivos H_" } }
        $contenido = (Get-Content $ultimo.FullName -Raw -ErrorAction Stop).Trim()
        $sec = $contenido.Substring($script:Config.PosicionSecuencia, $script:Config.LargoSecuencia).Trim()
        return @{ Secuencia = [int]$sec; Archivo = $ultimo.Name; Fecha = $ultimo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"); Error = $null }
    } catch {
        return @{ Secuencia = -1; Archivo = "ERROR"; Fecha = ""; Error = $_.Exception.Message }
    }
}

function Get-SecuenciaImpresa {
    try {
        $ultimoLog = Get-ChildItem -Path $script:Config.RutaLogs -Filter "log-*.txt" -ErrorAction Stop |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $ultimoLog) { return @{ Secuencia = -1; Log = "N/A"; Linea = ""; Error = "No se encontraron logs" } }
        $lineas = Get-Content $ultimoLog.FullName -Tail 500 -ErrorAction Stop
        $patron = "Fin impresion Kanban DOOR.*Secuencia\s+(\d+)"
        $ultimaSec = -1; $ultimaLinea = ""
        foreach ($l in $lineas) {
            if ($l -match $patron) { $ultimaSec = [int]$Matches[1]; $ultimaLinea = $l.Trim() }
        }
        if ($ultimaSec -eq -1) { return @{ Secuencia = -1; Log = $ultimoLog.Name; Linea = ""; Error = "No se encontro 'Fin impresion Kanban DOOR' en log" } }
        return @{ Secuencia = $ultimaSec; Log = $ultimoLog.Name; Linea = $ultimaLinea; Error = $null }
    } catch {
        return @{ Secuencia = -1; Log = "ERROR"; Linea = ""; Error = $_.Exception.Message }
    }
}

function Get-EstadoServicioRemoto {
    try {
        $svc = Get-Service -Name $script:Config.NombreServicio -ComputerName $script:Config.Servidor -ErrorAction Stop
        return @{ Status = $svc.Status.ToString(); Error = $null }
    } catch {
        return @{ Status = "Error"; Error = $_.Exception.Message }
    }
}

function Get-EstadoConsola {
    try {
        $proc = Get-WmiObject Win32_Process -ComputerName $script:Config.Servidor -Filter "Name='$($script:Config.ConsolaExe)'" -ErrorAction Stop
        if ($proc) { return @{ Running = $true; PID = $proc.ProcessId; Error = $null } }
        else { return @{ Running = $false; PID = $null; Error = $null } }
    } catch {
        return @{ Running = $false; PID = $null; Error = $_.Exception.Message }
    }
}

function Stop-ConsolaRemota {
    try {
        $proc = Get-WmiObject Win32_Process -ComputerName $script:Config.Servidor -Filter "Name='$($script:Config.ConsolaExe)'" -ErrorAction Stop
        if ($proc) {
            $proc.Terminate() | Out-Null
            return @{ Success = $true; Error = $null }
        }
        return @{ Success = $true; Error = "La consola no estaba corriendo" }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Start-ConsolaRemota {
    try {
        $result = ([wmiclass]"\\$($script:Config.Servidor)\root\cimv2:Win32_Process").Create($script:Config.ConsolaRuta)
        if ($result.ReturnValue -eq 0) { return @{ Success = $true; PID = $result.ProcessId; Error = $null } }
        else { return @{ Success = $false; PID = $null; Error = "ReturnValue: $($result.ReturnValue)" } }
    } catch {
        return @{ Success = $false; PID = $null; Error = $_.Exception.Message }
    }
}

function Stop-ServicioRemoto {
    try {
        $svc = Get-WmiObject Win32_Service -ComputerName $script:Config.Servidor -Filter "Name='$($script:Config.NombreServicio)'" -ErrorAction Stop
        if ($svc) { $svc.StopService() | Out-Null; return @{ Success = $true; Error = $null } }
        return @{ Success = $false; Error = "Servicio no encontrado" }
    } catch { return @{ Success = $false; Error = $_.Exception.Message } }
}

function Start-ServicioRemoto {
    try {
        $svc = Get-WmiObject Win32_Service -ComputerName $script:Config.Servidor -Filter "Name='$($script:Config.NombreServicio)'" -ErrorAction Stop
        if ($svc) { $svc.StartService() | Out-Null; return @{ Success = $true; Error = $null } }
        return @{ Success = $false; Error = "Servicio no encontrado" }
    } catch { return @{ Success = $false; Error = $_.Exception.Message } }
}

# ============================================================
# XAML - INTERFAZ GRAFICA
# ============================================================
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PrinterAlert - Monitor Kanban DOOR" Height="720" Width="860"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#0f1117" Foreground="White" FontFamily="Segoe UI">
    <Window.Resources>
        <Style x:Key="CardStyle" TargetType="Border">
            <Setter Property="Background" Value="#1a1d27"/>
            <Setter Property="CornerRadius" Value="12"/>
            <Setter Property="Padding" Value="20"/>
            <Setter Property="Margin" Value="6"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect BlurRadius="15" ShadowDepth="2" Opacity="0.3" Color="Black"/>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="BtnPrimary" TargetType="Button">
            <Setter Property="Background" Value="#3b82f6"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="18,10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#2563eb"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#374151"/>
                                <Setter Property="Foreground" Value="#6b7280"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="BtnDanger" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
            <Setter Property="Background" Value="#dc2626"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#b91c1c"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#374151"/>
                                <Setter Property="Foreground" Value="#6b7280"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="BtnSuccess" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
            <Setter Property="Background" Value="#16a34a"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#15803d"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#374151"/>
                                <Setter Property="Foreground" Value="#6b7280"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="LabelStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#9ca3af"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style x:Key="ValueStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="22"/>
            <Setter Property="FontWeight" Value="Bold"/>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <Border Grid.Row="0" Margin="6,0,6,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="&#x1F5A8;" FontSize="28" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <StackPanel>
                        <TextBlock Text="PrinterAlert" FontSize="24" FontWeight="Bold" Foreground="White"/>
                        <TextBlock Text="Monitor Kanban DOOR" FontSize="12" Foreground="#6b7280"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Modo:" Foreground="#9ca3af" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <ComboBox x:Name="cmbModo" Width="120" FontSize="13" Background="#252830" Foreground="Black" BorderBrush="#374151" Padding="6,4" SelectedIndex="0">
                        <ComboBoxItem Content="Demo"/>
                        <ComboBoxItem Content="Preguntar"/>
                        <ComboBoxItem Content="Auto"/>
                    </ComboBox>
                    <TextBlock Text="Tolerancia:" Foreground="#9ca3af" FontSize="13" VerticalAlignment="Center" Margin="16,0,8,0"/>
                    <TextBox x:Name="txtTolerancia" Text="5" Width="50" FontSize="14" FontWeight="Bold"
                             Background="#252830" Foreground="White" BorderBrush="#374151" BorderThickness="1"
                             Padding="8,6" HorizontalContentAlignment="Center" VerticalContentAlignment="Center"/>
                    <Button x:Name="btnRefresh" Content="&#x1F504; Actualizar" Style="{StaticResource BtnPrimary}" Margin="12,0,0,0"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- STATUS CARDS ROW -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <!-- Secuencia Recibida -->
            <Border Grid.Column="0" Style="{StaticResource CardStyle}">
                <StackPanel>
                    <TextBlock Text="SECUENCIA RECIBIDA" Style="{StaticResource LabelStyle}"/>
                    <TextBlock x:Name="lblSecRecibida" Text="---" Style="{StaticResource ValueStyle}" Foreground="#60a5fa"/>
                    <TextBlock x:Name="lblArchivoH" Text="" Style="{StaticResource LabelStyle}" FontSize="10" TextTrimming="CharacterEllipsis" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>
            <!-- Secuencia Impresa -->
            <Border Grid.Column="1" Style="{StaticResource CardStyle}">
                <StackPanel>
                    <TextBlock Text="SECUENCIA IMPRESA" Style="{StaticResource LabelStyle}"/>
                    <TextBlock x:Name="lblSecImpresa" Text="---" Style="{StaticResource ValueStyle}" Foreground="#a78bfa"/>
                    <TextBlock x:Name="lblArchivoLog" Text="" Style="{StaticResource LabelStyle}" FontSize="10" TextTrimming="CharacterEllipsis" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>
            <!-- Diferencia -->
            <Border Grid.Column="2" Style="{StaticResource CardStyle}" x:Name="cardDiferencia">
                <StackPanel>
                    <TextBlock Text="DIFERENCIA" Style="{StaticResource LabelStyle}"/>
                    <TextBlock x:Name="lblDiferencia" Text="---" Style="{StaticResource ValueStyle}"/>
                    <TextBlock x:Name="lblEstadoDif" Text="" Style="{StaticResource LabelStyle}" FontSize="10" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>
            <!-- Servicio -->
            <Border Grid.Column="3" Style="{StaticResource CardStyle}" x:Name="cardServicio">
                <StackPanel>
                    <TextBlock Text="SERVICIO" Style="{StaticResource LabelStyle}"/>
                    <TextBlock x:Name="lblServicio" Text="---" Style="{StaticResource ValueStyle}" FontSize="18"/>
                    <TextBlock x:Name="lblNombreServicio" Text="ImportTXTTASAsvcHost" Style="{StaticResource LabelStyle}" FontSize="10" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>
        </Grid>

        <!-- ALERTA -->
        <Border Grid.Row="2" x:Name="alertBanner" CornerRadius="10" Padding="16" Margin="6,8" Visibility="Collapsed">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="alertIcon" Text="&#x26A0;" FontSize="24" VerticalAlignment="Center" Margin="0,0,12,0"/>
                <StackPanel Grid.Column="1">
                    <TextBlock x:Name="alertTitle" FontSize="14" FontWeight="Bold" Foreground="White"/>
                    <TextBlock x:Name="alertMsg" FontSize="12" Foreground="#e5e7eb" TextWrapping="Wrap"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ACCIONES -->
        <Border Grid.Row="3" Style="{StaticResource CardStyle}">
            <StackPanel>
                <TextBlock Text="ACCIONES CORRECTIVAS" FontSize="13" FontWeight="SemiBold" Foreground="#9ca3af" Margin="0,0,0,12"/>
                <WrapPanel>
                    <Button x:Name="btnStopService" Content="&#x23F9; Detener Servicio" Style="{StaticResource BtnDanger}" Margin="0,0,8,8"/>
                    <Button x:Name="btnStartService" Content="&#x25B6; Iniciar Servicio" Style="{StaticResource BtnSuccess}" Margin="0,0,8,8"/>
                    <Button x:Name="btnStopConsola" Content="&#x23F9; Detener Consola" Style="{StaticResource BtnDanger}" Margin="0,0,8,8"/>
                    <Button x:Name="btnStartConsola" Content="&#x25B6; Iniciar Consola" Style="{StaticResource BtnSuccess}" Margin="0,0,8,8"/>
                    <Button x:Name="btnRunCorrectivo" Content="&#x1F527; Ejecutar Correctivo" Style="{StaticResource BtnPrimary}" Margin="0,0,8,8"/>
                    <Button x:Name="btnOpenLogs" Content="&#x1F4C2; Abrir Logs" Style="{StaticResource BtnPrimary}" Margin="0,0,8,8"/>
                    <Button x:Name="btnOpenProcessed" Content="&#x1F4C2; Abrir Procesados" Style="{StaticResource BtnPrimary}" Margin="0,0,8,8"/>
                </WrapPanel>
                <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                    <TextBlock Text="Consola:" Foreground="#9ca3af" FontSize="11" VerticalAlignment="Center" Margin="0,0,6,0"/>
                    <TextBlock x:Name="lblConsola" Text="---" Foreground="#6b7280" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <!-- LOG -->
        <Border Grid.Row="4" Style="{StaticResource CardStyle}">
            <DockPanel>
                <TextBlock DockPanel.Dock="Top" Text="REGISTRO DE ACTIVIDAD" FontSize="13" FontWeight="SemiBold" Foreground="#9ca3af" Margin="0,0,0,8"/>
                <TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                         Background="#12141c" Foreground="#d1d5db" BorderThickness="0" FontFamily="Consolas" FontSize="11"
                         Padding="10"/>
            </DockPanel>
        </Border>

        <!-- FOOTER -->
        <Border Grid.Row="5" Margin="6,4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="lblUltimaAct" Text="Sin actualizar" Foreground="#6b7280" FontSize="11" VerticalAlignment="Center"/>
                <CheckBox x:Name="chkAutoRefresh" Grid.Column="1" Content="Auto-refresh (30s)" Foreground="#9ca3af" FontSize="11" VerticalAlignment="Center" Margin="0,0,12,0"/>
                <TextBlock Grid.Column="2" Text="Servidor: 172.17.132.153" Foreground="#4b5563" FontSize="11" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ============================================================
# CREAR VENTANA
# ============================================================
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Obtener controles
$lblSecRecibida    = $window.FindName("lblSecRecibida")
$lblSecImpresa     = $window.FindName("lblSecImpresa")
$lblDiferencia     = $window.FindName("lblDiferencia")
$lblEstadoDif      = $window.FindName("lblEstadoDif")
$lblServicio       = $window.FindName("lblServicio")
$lblArchivoH       = $window.FindName("lblArchivoH")
$lblArchivoLog     = $window.FindName("lblArchivoLog")
$lblUltimaAct      = $window.FindName("lblUltimaAct")
$txtTolerancia     = $window.FindName("txtTolerancia")
$txtLog            = $window.FindName("txtLog")
$btnRefresh        = $window.FindName("btnRefresh")
$btnStopService    = $window.FindName("btnStopService")
$btnStartService   = $window.FindName("btnStartService")
$btnStopConsola    = $window.FindName("btnStopConsola")
$btnStartConsola   = $window.FindName("btnStartConsola")
$btnRunCorrectivo  = $window.FindName("btnRunCorrectivo")
$btnOpenLogs       = $window.FindName("btnOpenLogs")
$btnOpenProcessed  = $window.FindName("btnOpenProcessed")
$chkAutoRefresh    = $window.FindName("chkAutoRefresh")
$cmbModo           = $window.FindName("cmbModo")
$lblConsola        = $window.FindName("lblConsola")
$alertBanner       = $window.FindName("alertBanner")
$alertTitle        = $window.FindName("alertTitle")
$alertMsg          = $window.FindName("alertMsg")
$alertIcon         = $window.FindName("alertIcon")
$cardDiferencia    = $window.FindName("cardDiferencia")

# ============================================================
# FUNCIONES UI
# ============================================================
function Add-Log {
    param([string]$Mensaje, [string]$Tipo = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$timestamp] [$Tipo] $Mensaje`r`n")
    $txtLog.ScrollToEnd()
}

function Update-Dashboard {
    Add-Log "Actualizando dashboard..." "INFO"

    # Tolerancia
    $tol = 5
    if ([int]::TryParse($txtTolerancia.Text, [ref]$tol)) {
        $script:Config.Tolerancia = $tol
    } else {
        $tol = $script:Config.Tolerancia
    }

    # Secuencia recibida
    $rec = Get-SecuenciaRecibida
    if ($rec.Error) {
        $lblSecRecibida.Text = "ERR"
        $lblArchivoH.Text = $rec.Error
        Add-Log "Error leyendo H_: $($rec.Error)" "ERROR"
    } else {
        $lblSecRecibida.Text = $rec.Secuencia.ToString()
        $lblArchivoH.Text = "$($rec.Archivo) ($($rec.Fecha))"
        Add-Log "Secuencia recibida: $($rec.Secuencia) ($($rec.Archivo))"
    }

    # Secuencia impresa
    $imp = Get-SecuenciaImpresa
    if ($imp.Error) {
        $lblSecImpresa.Text = "ERR"
        $lblArchivoLog.Text = $imp.Error
        Add-Log "Error leyendo log: $($imp.Error)" "ERROR"
    } else {
        $lblSecImpresa.Text = $imp.Secuencia.ToString()
        $lblArchivoLog.Text = $imp.Log
        Add-Log "Secuencia impresa: $($imp.Secuencia) ($($imp.Log))"
    }

    # Diferencia
    if ($rec.Secuencia -ge 0 -and $imp.Secuencia -ge 0) {
        $dif = $rec.Secuencia - $imp.Secuencia
        $lblDiferencia.Text = $dif.ToString()

        if ($dif -le 0) {
            $lblDiferencia.Foreground = [System.Windows.Media.Brushes]::LimeGreen
            $lblEstadoDif.Text = "AL DIA"
            $alertBanner.Visibility = "Collapsed"
        } elseif ($dif -le $tol) {
            $lblDiferencia.Foreground = [System.Windows.Media.Brushes]::Orange
            $lblEstadoDif.Text = "DENTRO DE TOLERANCIA ($tol)"
            $alertBanner.Visibility = "Visible"
            $alertBanner.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#78592d1a")
            $alertBanner.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ca8a04")
            $alertBanner.BorderThickness = [System.Windows.Thickness]::new(1)
            $alertTitle.Text = "Atencion: Diferencia dentro de tolerancia"
            $alertMsg.Text = "Hay $dif secuencia(s) pendientes de impresion. Tolerancia: $tol"
            $alertIcon.Text = [char]0x26A0
            Add-Log "Diferencia $dif dentro de tolerancia ($tol)" "WARN"
        } else {
            $lblDiferencia.Foreground = [System.Windows.Media.Brushes]::Red
            $lblEstadoDif.Text = "SUPERA TOLERANCIA ($tol)"
            $alertBanner.Visibility = "Visible"
            $alertBanner.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#78450a0a")
            $alertBanner.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#dc2626")
            $alertBanner.BorderThickness = [System.Windows.Thickness]::new(1)
            $alertTitle.Text = "ALERTA: La impresion esta atrasada!"
            $alertMsg.Text = "Diferencia de $dif supera la tolerancia de $tol. Se recomienda: Detener servicio, reiniciar consola, esperar impresion y reiniciar servicio."
            $alertIcon.Text = [char]0x274C
            Add-Log "ALERTA! Diferencia $dif SUPERA tolerancia ($tol)" "ALERTA"
        }
    }

    # Servicio
    $svc = Get-EstadoServicioRemoto
    if ($svc.Error) {
        $lblServicio.Text = "Error"
        $lblServicio.Foreground = [System.Windows.Media.Brushes]::Red
        Add-Log "Error consultando servicio: $($svc.Error)" "ERROR"
    } elseif ($svc.Status -eq "Running") {
        $lblServicio.Text = "Running"
        $lblServicio.Foreground = [System.Windows.Media.Brushes]::LimeGreen
        Add-Log "Servicio: Running"
    } else {
        $lblServicio.Text = $svc.Status
        $lblServicio.Foreground = [System.Windows.Media.Brushes]::Red
        Add-Log "Servicio: $($svc.Status)" "WARN"
    }

    # Consola
    $con = Get-EstadoConsola
    if ($con.Running) {
        $lblConsola.Text = "Running (PID: $($con.PID))"
        $lblConsola.Foreground = [System.Windows.Media.Brushes]::LimeGreen
    } else {
        $lblConsola.Text = "Detenida"
        $lblConsola.Foreground = [System.Windows.Media.Brushes]::Red
    }

    $lblUltimaAct.Text = "Ultima actualizacion: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

# ============================================================
# EVENTOS
# ============================================================
$btnRefresh.Add_Click({ Update-Dashboard })

$btnStopService.Add_Click({
    $modo = $cmbModo.SelectedIndex  # 0=Demo, 1=Preguntar, 2=Auto
    if ($modo -eq 0) {
        Add-Log "[DEMO] Se detendria el servicio $($script:Config.NombreServicio)" "DEMO"
        [System.Windows.MessageBox]::Show("MODO DEMO: Se detendria el servicio $($script:Config.NombreServicio).`nCambie a modo 'Preguntar' o 'Auto' para ejecutar.", "Modo Demo", "OK", "Information")
        return
    }
    if ($modo -eq 1) {
        $result = [System.Windows.MessageBox]::Show("Detener servicio $($script:Config.NombreServicio)?", "Confirmar", "YesNo", "Warning")
        if ($result -ne "Yes") { Add-Log "Detencion de servicio cancelada por usuario" "INFO"; return }
    }
    Add-Log "Deteniendo servicio $($script:Config.NombreServicio)..." "ACCION"
    $r = Stop-ServicioRemoto
    if ($r.Success) { Add-Log "Servicio detenido exitosamente." "OK" }
    else { Add-Log "Error deteniendo servicio: $($r.Error)" "ERROR" }
    Start-Sleep -Seconds 2; Update-Dashboard
})

$btnStartService.Add_Click({
    $modo = $cmbModo.SelectedIndex
    if ($modo -eq 0) {
        Add-Log "[DEMO] Se iniciaria el servicio $($script:Config.NombreServicio)" "DEMO"
        [System.Windows.MessageBox]::Show("MODO DEMO: Se iniciaria el servicio.`nCambie a modo 'Preguntar' o 'Auto' para ejecutar.", "Modo Demo", "OK", "Information")
        return
    }
    if ($modo -eq 1) {
        $result = [System.Windows.MessageBox]::Show("Iniciar servicio $($script:Config.NombreServicio)?", "Confirmar", "YesNo", "Question")
        if ($result -ne "Yes") { Add-Log "Inicio de servicio cancelado por usuario" "INFO"; return }
    }
    Add-Log "Iniciando servicio $($script:Config.NombreServicio)..." "ACCION"
    $r = Start-ServicioRemoto
    if ($r.Success) { Add-Log "Servicio iniciado exitosamente." "OK" }
    else { Add-Log "Error iniciando servicio: $($r.Error)" "ERROR" }
    Start-Sleep -Seconds 2; Update-Dashboard
})

$btnStopConsola.Add_Click({
    $modo = $cmbModo.SelectedIndex
    if ($modo -eq 0) {
        Add-Log "[DEMO] Se detendria la consola $($script:Config.ConsolaExe)" "DEMO"
        [System.Windows.MessageBox]::Show("MODO DEMO: Se detendria la consola de impresion.`nCambie a modo 'Preguntar' o 'Auto' para ejecutar.", "Modo Demo", "OK", "Information")
        return
    }
    if ($modo -eq 1) {
        $result = [System.Windows.MessageBox]::Show("Detener la consola de impresion?", "Confirmar", "YesNo", "Warning")
        if ($result -ne "Yes") { Add-Log "Detencion de consola cancelada por usuario" "INFO"; return }
    }
    Add-Log "Deteniendo consola de impresion..." "ACCION"
    $r = Stop-ConsolaRemota
    if ($r.Success) { Add-Log "Consola detenida. $($r.Error)" "OK" }
    else { Add-Log "Error deteniendo consola: $($r.Error)" "ERROR" }
    Start-Sleep -Seconds 2; Update-Dashboard
})

$btnStartConsola.Add_Click({
    $modo = $cmbModo.SelectedIndex
    if ($modo -eq 0) {
        Add-Log "[DEMO] Se iniciaria la consola $($script:Config.ConsolaExe)" "DEMO"
        [System.Windows.MessageBox]::Show("MODO DEMO: Se iniciaria la consola de impresion.`nCambie a modo 'Preguntar' o 'Auto' para ejecutar.", "Modo Demo", "OK", "Information")
        return
    }
    if ($modo -eq 1) {
        $result = [System.Windows.MessageBox]::Show("Iniciar la consola de impresion?", "Confirmar", "YesNo", "Question")
        if ($result -ne "Yes") { Add-Log "Inicio de consola cancelado por usuario" "INFO"; return }
    }
    Add-Log "Iniciando consola de impresion..." "ACCION"
    $r = Start-ConsolaRemota
    if ($r.Success) { Add-Log "Consola iniciada (PID: $($r.PID))" "OK" }
    else { Add-Log "Error iniciando consola: $($r.Error)" "ERROR" }
    Start-Sleep -Seconds 3; Update-Dashboard
})

$btnRunCorrectivo.Add_Click({
    $modo = $cmbModo.SelectedIndex
    $pasos = "1. Detener servicio ImportTXTTASAsvcHost`n2. Detener consola de impresion`n3. Iniciar consola de impresion`n4. Esperar impresion de ultima secuencia`n5. Iniciar servicio ImportTXTTASAsvcHost"
    if ($modo -eq 0) {
        Add-Log "[DEMO] Se ejecutaria el correctivo completo" "DEMO"
        [System.Windows.MessageBox]::Show("MODO DEMO - Pasos que se ejecutarian:`n`n$pasos`n`nCambie a modo 'Preguntar' o 'Auto' para ejecutar.", "Modo Demo", "OK", "Information")
        return
    }
    if ($modo -eq 1) {
        $result = [System.Windows.MessageBox]::Show("Ejecutar correctivo completo?`n`n$pasos", "Confirmar correctivo", "YesNo", "Warning")
        if ($result -ne "Yes") { Add-Log "Correctivo cancelado por usuario" "INFO"; return }
    }
    # Paso 1: Detener servicio
    Add-Log "CORRECTIVO PASO 1/5: Deteniendo servicio..." "ACCION"
    $r = Stop-ServicioRemoto
    if ($r.Success) { Add-Log "Servicio detenido." "OK" } else { Add-Log "Error: $($r.Error)" "ERROR" }
    Start-Sleep -Seconds 2
    # Paso 2: Detener consola
    Add-Log "CORRECTIVO PASO 2/5: Deteniendo consola..." "ACCION"
    $r = Stop-ConsolaRemota
    if ($r.Success) { Add-Log "Consola detenida." "OK" } else { Add-Log "Error: $($r.Error)" "ERROR" }
    Start-Sleep -Seconds 3
    # Paso 3: Iniciar consola
    Add-Log "CORRECTIVO PASO 3/5: Iniciando consola..." "ACCION"
    $r = Start-ConsolaRemota
    if ($r.Success) { Add-Log "Consola iniciada (PID: $($r.PID))" "OK" } else { Add-Log "Error: $($r.Error)" "ERROR" }
    Start-Sleep -Seconds 5
    # Paso 4: Esperar impresion
    Add-Log "CORRECTIVO PASO 4/5: Esperando impresion de ultima secuencia..." "ACCION"
    $recActual = Get-SecuenciaRecibida
    $maxWait = 30; $waited = 0
    while ($waited -lt $maxWait) {
        $impActual = Get-SecuenciaImpresa
        if ($null -ne $impActual -and $impActual.Secuencia -ge $recActual.Secuencia) {
            Add-Log "Secuencia $($impActual.Secuencia) impresa OK!" "OK"
            break
        }
        $waited++
        Add-Log "Esperando... ($waited/$maxWait) Impresa: $($impActual.Secuencia) / Esperada: $($recActual.Secuencia)" "INFO"
        Start-Sleep -Seconds 10
    }
    if ($waited -ge $maxWait) { Add-Log "TIMEOUT esperando impresion" "ERROR" }
    # Paso 5: Iniciar servicio
    Add-Log "CORRECTIVO PASO 5/5: Iniciando servicio..." "ACCION"
    $r = Start-ServicioRemoto
    if ($r.Success) { Add-Log "Servicio iniciado. CORRECTIVO COMPLETO." "OK" } else { Add-Log "Error: $($r.Error)" "ERROR" }
    Update-Dashboard
})

$btnOpenLogs.Add_Click({
    try { Start-Process "explorer.exe" -ArgumentList $script:Config.RutaLogs }
    catch { Add-Log "Error abriendo carpeta de logs: $_" "ERROR" }
})

$btnOpenProcessed.Add_Click({
    try { Start-Process "explorer.exe" -ArgumentList $script:Config.RutaProcesados }
    catch { Add-Log "Error abriendo carpeta de procesados: $_" "ERROR" }
})

# Auto-refresh timer
$script:timer = New-Object System.Windows.Threading.DispatcherTimer
$script:timer.Interval = [TimeSpan]::FromSeconds(30)
$script:timer.Add_Tick({ Update-Dashboard })

$chkAutoRefresh.Add_Checked({ $script:timer.Start(); Add-Log "Auto-refresh activado (30s)" "INFO" })
$chkAutoRefresh.Add_Unchecked({ $script:timer.Stop(); Add-Log "Auto-refresh desactivado" "INFO" })

# Cargar al iniciar
$window.Add_Loaded({ Update-Dashboard })

# ============================================================
# MOSTRAR VENTANA
# ============================================================
$window.ShowDialog() | Out-Null
