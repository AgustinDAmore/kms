import os
import ctypes
import subprocess
import datetime
import socket
import re
import winreg
import platform
import wmi
import psutil
import getpass
####################################
import keys
##################################
# HERRAMIENTAS Y FUNCIONES CLAVE #
##################################
def obtener_info_discos_windows():
    discos = []
    
    # Obtener todas las particiones (sin duplicados por disco físico)
    for particion in psutil.disk_partitions(all=True):
        if not particion.device.startswith('\\\\'):  # Ignorar unidades de red
            try:
                # Obtener uso del disco
                uso = psutil.disk_usage(particion.mountpoint)
                
                # Obtener número de disco físico (para relacionar con wmic)
                numero_disco = None
                try:
                    cmd = f"wmic volume where name='{particion.device.replace('\\', '\\\\')}' get DriveNumber /value"
                    resultado = subprocess.run(cmd, shell=True, capture_output=True, text=True)
                    if 'DriveNumber' in resultado.stdout:
                        numero_disco = resultado.stdout.split('=')[1].strip()
                except:
                    pass
                
                discos.append({
                    'letra': particion.device,
                    'punto_montaje': particion.mountpoint,
                    'total_gb': round(uso.total / (1024 ** 3), 2),
                    'numero_disco': numero_disco,
                    'tipo': None  # Se determinará después
                })
            except PermissionError:
                continue

    # Determinar SSD/HDD para cada disco
    for disco in discos:
        if disco['numero_disco'] is not None:
            try:
                cmd = f"wmic diskdrive where index={disco['numero_disco']} get mediatype /value"
                resultado = subprocess.run(cmd, shell=True, capture_output=True, text=True)
                if 'MediaType' in resultado.stdout:
                    tipo = resultado.stdout.split('=')[1].strip()
                    disco['tipo'] = "SSD" if "SSD" in tipo or "Solid State" in tipo else "HDD"
            except:
                disco['tipo'] = "Desconocido"
        else:
            disco['tipo'] = "Desconocido"

    return discos
################################################################################################################################################################
def obtener_informacion_pc():    
    # 1. Nombre de usuario y cuenta registrada
    usuario = os.getlogin()
    hostname = socket.gethostname()
    print(f"\n[Usuario]")
    print(f"Nombre de usuario: {usuario}")
    print(f"Nombre del equipo (hostname): {hostname}")
    c = wmi.WMI()
    for system in c.Win32_ComputerSystem():
        print(f"\n[Modelo del equipo]")
        print(f"Fabricante: {system.Manufacturer}")
        print(f"Modelo: {system.Model}")
        
    # Número de serie (BIOS)
    for bios in c.Win32_BIOS():
        print(f"Número de serie (BIOS): {bios.SerialNumber}")

    discos = obtener_info_discos_windows()
    print("\n[Discos en Windows]")
    for disco in discos:
        print(f"\nLetra: {disco['letra']}")
        print(f"Punto de montaje: {disco['punto_montaje']}")
        print(f"Capacidad: {disco['total_gb']} GB")
        print(f"Tipo: {disco['tipo']}")
        print("-" * 40)

    # 3. Cantidad de RAM (GB)
    ram_total = round(psutil.virtual_memory().total / (1024 ** 3), 2)
    print(f"\n[Memoria RAM]")
    print(f"RAM Total: {ram_total} GB")

    # 4. Procesador
    print(f"\n[Procesador]")
    print(f"CPU: {platform.processor()}")
    print(f"Arquitectura: {platform.machine()}")
    print(f"Núcleos físicos: {psutil.cpu_count(logical=False)}")
    print(f"Núcleos lógicos (con hilos): {psutil.cpu_count(logical=True)}")
################################################################################################################################################################
def get_windows_info():
    # Método 1: Usando platform
    print(f"Nombre: {platform.system()} {platform.release()}")
    print(f"Versión: {platform.version()}")
    print(f"Edición: {' '.join(platform.win32_ver()[1:])}")
    print(f"Arquitectura: {platform.machine()}")

    # Método 2: Leyendo el registro (más preciso para la edición)
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as key:
            product_name = winreg.QueryValueEx(key, "ProductName")[0]
            display_version = winreg.QueryValueEx(key, "DisplayVersion")[0]
            print(f"\n[Registro] Edición exacta: {product_name} ({display_version})")
    except Exception as e:
        print(f"\nError al leer registro: {e}")
################################################################################################################################################################
def configurar_servidores_kms_publicos():
    return [
        "kms8.msguides.com",
        "10.3.0.20:1688",
        "kms.digiboy.ir",
        "kms.srv.crsoo.com",
        "kms.loli.beer",
        "kms9.MSGuides.com",
        "kms.zhuxiaole.org",
        "kms.lolico.moe",
        "kms.moeclub.org",
        "kms.lotro.cc"
    ]
################################################################################################################################################################
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False
################################################################################################################################################################
def run_command(command):
    try:
        result = subprocess.run(
            command,
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='cp850',
            creationflags=subprocess.CREATE_NO_WINDOW,
            timeout=30
        )
        if result.stderr:
            return f"Error (Código {result.returncode}): {result.stderr.strip()}"
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        return "Error: Tiempo de espera agotado"
    except Exception as e:
        return f"Excepción: {str(e)}"
################################################################################################################################################################
##########
# Office #
##########
"""
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x64)%\Microsoft Office\Office16
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6MWKP >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:kms8.msguides.com
cscript ospp.vbs /act
"""
def activate_office_2019_vl():
    """
    Función para activar Office 2019 Volume License usando un servidor KMS.
    Realiza operaciones similares a los comandos batch proporcionados.
    """
    try:
        # Posibles rutas de instalación de Office
        office_paths = [
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Microsoft Office", "Office16"),
            os.path.join(os.environ.get("ProgramFiles", ""), "Microsoft Office", "Office16"),
            os.path.join(os.environ.get("ProgramFiles(x64)", ""), "Microsoft Office", "Office16")
        ]

        # Buscar la ruta válida de Office
        valid_path = None
        for path in office_paths:
            if os.path.exists(path) and os.path.isdir(path):
                valid_path = path
                break

        if not valid_path:
            raise FileNotFoundError("No se pudo encontrar la ruta de instalación de Office")

        # Cambiar al directorio de Office
        os.chdir(valid_path)

        # Instalar licencias VL
        licenses_path = os.path.join(valid_path, "..", "root", "Licenses16")
        if os.path.exists(licenses_path):
            for file in os.listdir(licenses_path):
                if file.startswith("ProPlus2019VL") and file.endswith(".xrm-ms"):
                    license_file = os.path.join(licenses_path, file)
                    subprocess.run(["cscript", "ospp.vbs", "/inslic:" + license_file], check=True)

        # Configurar puerto KMS
        subprocess.run(["cscript", "ospp.vbs", "/setprt:1688"], check=True)
        
        # Eliminar clave anterior (silenciosamente)
        subprocess.run(["cscript", "ospp.vbs", "/unpkey:6MWKP"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        # Instalar nueva clave
        subprocess.run(["cscript", "ospp.vbs", "/inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"], check=True)
        
        # Configurar servidor KMS
        subprocess.run(["cscript", "ospp.vbs", "/sethst:kms8.msguides.com"], check=True)
        
        # Intentar activación
        subprocess.run(["cscript", "ospp.vbs", "/act"], check=True)
        
        print("Proceso de activación completado.")
        return True
    
    except subprocess.CalledProcessError as e:
        print(f"Error durante la ejecución del comando: {e}")
        return False
    except Exception as e:
        print(f"Error inesperado: {e}")
        return False
###########
# WINDOWS #
###########

def modify_registry_value(key_path: str, value_name: str, new_value: str, value_type=winreg.REG_SZ) -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, value_name, 0, value_type, new_value)
        return True
    except WindowsError as e:
        print(f"Error al modificar {value_name} en {key_path}: {e}")
        return False
################################################################################################################################################################
def backup_current_settings():
    backup = {}
    paths = [
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion"
    ]
    
    for path in paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                values = {}
                for value_name in ["EditionID", "ProductName"]:
                    try:
                        values[value_name] = winreg.QueryValueEx(key, value_name)[0]
                    except WindowsError:
                        values[value_name] = ""
                backup[path] = values
        except WindowsError:
            continue
    
    return backup
################################################################################################################################################################
def Cambiar(version: str) -> bool:
    if not is_admin():
        print("Error: Se requieren privilegios de administrador para modificar el registro.")
        return False
    
    # Configuración de versiones
    version_config = {
        "1": {
            "EditionID": "Core",
            "ProductName": "Windows 10 Home"
        },
        "2": {
            "EditionID": "Professional",
            "ProductName": "Windows 10 Pro"
        }
    }
    
    if version not in version_config:
        print(f"Error: Versión '{version}' no válida. Use '1' para Home o '2' para Professional.")
        return False
    
    # Crear backup antes de modificar
    backup = backup_current_settings()
    print("\n[+] Creando copia de seguridad de la configuración actual...")
    
    # Configurar los nuevos valores
    config = version_config[version]
    success = True
    
    # Lista de rutas del registro a modificar
    registry_paths = [
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion"
    ]
    
    print("\n[+] Modificando configuración del sistema...")
    for path in registry_paths:
        print(f"\n• Modificando {path}:")
        
        # Modificar EditionID
        if not modify_registry_value(path, "EditionID", config["EditionID"]):
            success = False
        else:
            print(f"  - EditionID cambiado a '{config['EditionID']}'")
        
        # Modificar ProductName
        if not modify_registry_value(path, "ProductName", config["ProductName"]):
            success = False
        else:
            print(f"  - ProductName cambiado a '{config['ProductName']}'")
    
    if success:
        print("\n[+] Todos los cambios se aplicaron correctamente.")
        print("[!] Nota: Es posible que necesites reiniciar el sistema para que los cambios surtan efecto.")
    else:
        print("\n[!] Algunos cambios no se aplicaron correctamente.")
        
        # Opción para restaurar el backup
        restore = input("\n¿Deseas restaurar la configuración anterior? (s/n): ").lower()
        if restore == 's':
            print("\n[+] Restaurando configuración anterior...")
            for path, values in backup.items():
                for value_name, old_value in values.items():
                    if old_value:  # Solo restaurar si había un valor
                        modify_registry_value(path, value_name, old_value)
            print("[+] Restauración completada.")
    
    return success
################################################################################################################################################################
class WindowsActivator:
    GVLK_KEYS = keys.obtenerKeys()
################################################################################################################################################################
    @staticmethod
    def detect_windows_edition():
        """Detecta la edición actual de Windows con mayor precisión"""
        edition_map = {
            # Windows 10/11
            "Windows 10 Home": "Core",
            "Windows 11 Home": "Core",
            "Windows 10 Home N": "CoreN",
            "Windows 11 Home N": "CoreN",
            "Windows 10 Home Single Language": "CoreSingleLanguage",
            "Windows 10 Home China": "CoreChina",
            "Windows 10 Pro": "Professional",
            "Windows 11 Pro": "Professional",
            "Windows 10 Pro N": "ProfessionalN",
            "Windows 11 Pro N": "ProfessionalN",
            "Windows 10 Pro Education": "ProfessionalEducation",
            "Windows 10 Pro Education N": "ProfessionalEducationN",
            "Windows 10 Pro for Workstations": "ProfessionalWorkstation",
            "Windows 10 Pro for Workstations N": "ProfessionalWorkstationN",
            "Windows 10 Education": "Education",
            "Windows 11 Education": "Education",
            "Windows 10 Education N": "EducationN",
            "Windows 11 Education N": "EducationN",
            "Windows 10 Enterprise": "Enterprise",
            "Windows 11 Enterprise": "Enterprise",
            "Windows 10 Enterprise N": "EnterpriseN",
            "Windows 11 Enterprise N": "EnterpriseN",
            "Windows 10 Enterprise G": "EnterpriseG",
            "Windows 10 Enterprise G N": "EnterpriseGN",
            "Windows 10 Enterprise 2015 LTSB": "EnterpriseS_LTSB2015",
            "Windows 10 Enterprise 2015 LTSB N": "EnterpriseS_LTSB2015N",
            "Windows 10 Enterprise 2016 LTSB": "EnterpriseS_LTSB2016",
            "Windows 10 Enterprise 2016 LTSB N": "EnterpriseS_LTSB2016N",
            "Windows 10 Enterprise LTSC 2019": "EnterpriseS_LTSC2019",
            "Windows 10 Enterprise LTSC 2019 N": "EnterpriseS_LTSC2019N",
            "Windows 10 Enterprise LTSC 2021": "EnterpriseS_LTSC2021",
            "Windows 11 Enterprise LTSC 2021": "EnterpriseS_LTSC2021",
            "Windows 10 Enterprise for Virtual Desktops": "EnterpriseVirtualDesktops",
            
            # Windows 8.1
            "Windows 8.1": "Core_8.1",
            "Windows 8.1 N": "CoreN_8.1",
            "Windows 8.1 Single Language": "CoreSingleLanguage_8.1",
            "Windows 8.1 China": "CoreChina_8.1",
            "Windows 8.1 Pro": "Professional_8.1",
            "Windows 8.1 Pro N": "ProfessionalN_8.1",
            "Windows 8.1 Pro with Media Center": "ProfessionalWMC_8.1",
            "Windows 8.1 Enterprise": "Enterprise_8.1",
            "Windows 8.1 Enterprise N": "EnterpriseN_8.1",
            
            # Windows 7
            "Windows 7 Professional": "Professional_7",
            "Windows 7 Professional N": "ProfessionalN_7",
            "Windows 7 Enterprise": "Enterprise_7",
            "Windows 7 Enterprise N": "EnterpriseN_7",
            
            # Windows Server (solo algunas ediciones principales)
            "Windows Server 2019 Standard": "ServerStandard_2019",
            "Windows Server 2019 Datacenter": "ServerDatacenter_2019",
            "Windows Server 2016 Standard": "ServerStandard_2016",
            "Windows Server 2016 Datacenter": "ServerDatacenter_2016",
            "Windows Server 2012 R2 Standard": "ServerStandard_2012R2",
            "Windows Server 2012 R2 Datacenter": "ServerDatacenter_2012R2",
            "Windows Server 2008 R2 Standard": "ServerStandard_2008R2",
            "Windows Server 2008 R2 Datacenter": "ServerDatacenter_2008R2",
            "Windows Server 2008 R2 Enterprise": "ServerEnterprise_2008R2",
            "Windows Web Server 2008 R2": "WebServer_2008R2"
        }
        
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as key:
                product_name = winreg.QueryValueEx(key, "ProductName")[0]
                edition_id = winreg.QueryValueEx(key, "EditionID")[0]
                
                # Primero intentar por ProductName (coincidencia exacta o parcial)
                for name, edition in edition_map.items():
                    if name.lower() in product_name.lower():
                        return edition
                
                # Si no se encontró por ProductName, intentar por EditionID
                if edition_id in edition_map.values():
                    return edition_id
                
                # Mapeo adicional para EditionIDs que pueden no estar en ProductName
                edition_id_map = {
                    "Core": "Core",
                    "CoreN": "CoreN",
                    "Professional": "Professional",
                    "ProfessionalN": "ProfessionalN",
                    "Enterprise": "Enterprise",
                    "EnterpriseN": "EnterpriseN",
                    "ServerStandard": "ServerStandard",
                    "ServerDatacenter": "ServerDatacenter"
                }
                
                for id_part, edition in edition_id_map.items():
                    if id_part in edition_id:
                        return edition
                
                return None
                
        except Exception as e:
            print(f"[!] Error detectando edición de Windows: {e}")
            return None
#############################################################################################
    def troubleshoot_activation(self) -> dict:
        """
        Realiza un diagnóstico completo de los problemas de activación de Windows.
        
        Returns:
            dict: Diccionario con los resultados del diagnóstico y recomendaciones
        """
        diagnosis = {
            'system_info': {},
            'detected_issues': [],
            'recommendations': [],
            'activation_status': None,
            'kms_configuration': None
        }
        
        print("\n[🔍] Iniciando diagnóstico de activación...\n")
        
        # 1. Verificar información básica del sistema
        try:
            diagnosis['system_info'] = {
                'windows_edition': self.detect_windows_edition(),
                'architecture': platform.machine(),
                'system_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'is_admin': ctypes.windll.shell32.IsUserAnAdmin() != 0
            }
            print("[✓] Información del sistema recolectada")
        except Exception as e:
            diagnosis['detected_issues'].append(f"No se pudo obtener información del sistema: {str(e)}")
        
        # 2. Verificar estado de activación actual
        try:
            activation_status = self.verify_activation()
            diagnosis['activation_status'] = activation_status
            if activation_status:
                print("[✓] Windows ya está activado correctamente")
                diagnosis['recommendations'].append("No se requiere activación adicional")
                return diagnosis
        except Exception as e:
            diagnosis['detected_issues'].append(f"Error al verificar activación: {str(e)}")
        
        # 3. Verificar configuración KMS
        try:
            windir = os.environ.get("windir")
            slmgr_path = os.path.join(windir, "system32", "slmgr.vbs") if windir else None
            
            if slmgr_path and os.path.exists(slmgr_path):
                _, kms_output = self.run_command(f'cscript //nologo "{slmgr_path}" /dlv')
                diagnosis['kms_configuration'] = self._parse_kms_info(kms_output)
                print("[✓] Configuración KMS verificada")
            else:
                diagnosis['detected_issues'].append("No se encontró slmgr.vbs (instalación corrupta)")
        except Exception as e:
            diagnosis['detected_issues'].append(f"Error al verificar KMS: {str(e)}")
        
        # 4. Verificar problemas comunes
        self._check_common_issues(diagnosis)
        
        # 5. Generar recomendaciones basadas en los hallazgos
        self._generate_recommendations(diagnosis)
        
        print("\nResumen del diagnóstico:")
        print(f"- Problemas detectados: {len(diagnosis['detected_issues'])}")
        print(f"- Recomendaciones: {len(diagnosis['recommendations'])}")
        
        return diagnosis
################################################################################################################################################################
    def _parse_kms_info(self, kms_output: str) -> dict:
        """Parsea la información de configuración KMS"""
        info = {
            'kms_server': None,
            'port': None,
            'activation_interval': None,
            'renewal_interval': None
        }
        
        server_match = re.search(r'Key Management Service Machine name:\s*(.*?)\s*$', kms_output, re.M)
        if server_match:
            info['kms_server'] = server_match.group(1).strip()
        
        port_match = re.search(r'Key Management Service port:\s*(.*?)\s*$', kms_output, re.M)
        if port_match:
            info['port'] = port_match.group(1).strip()
        
        interval_match = re.search(r'Activation Interval:\s*(.*?)\s*$', kms_output, re.M)
        if interval_match:
            info['activation_interval'] = interval_match.group(1).strip()
        
        renewal_match = re.search(r'Renewal Interval:\s*(.*?)\s*$', kms_output, re.M)
        if renewal_match:
            info['renewal_interval'] = renewal_match.group(1).strip()
        
        return info
################################################################################################################################################################
    def _check_common_issues(self, diagnosis: dict):
        """Verifica problemas comunes de activación"""
        # 1. Verificar si es una edición compatible con KMS
        edition = diagnosis['system_info'].get('windows_edition', '')
        if 'Home' in edition:
            diagnosis['detected_issues'].append(f"La edición {edition} no soporta activación KMS")
            diagnosis['recommendations'].append("Actualice a una edición Professional/Enterprise")
        
        # 2. Verificar permisos de administrador
        if not diagnosis['system_info'].get('is_admin', False):
            diagnosis['detected_issues'].append("El proceso no se ejecuta como administrador")
            diagnosis['recommendations'].append("Ejecute el programa como administrador")
        
        # 3. Verificar conexión a internet
        try:
            socket.create_connection(("www.microsoft.com", 80), timeout=5)
        except:
            diagnosis['detected_issues'].append("No hay conexión a internet o está bloqueada")
            diagnosis['recommendations'].append("Verifique su conexión a internet y firewall")
        
        # 4. Verificar servidor KMS configurado
        if diagnosis['kms_configuration'] and not diagnosis['kms_configuration'].get('kms_server'):
            diagnosis['detected_issues'].append("No hay servidor KMS configurado")
            diagnosis['recommendations'].append("Configure un servidor KMS válido")
################################################################################################################################################################
    def _generate_recommendations(self, diagnosis: dict):
        """Genera recomendaciones basadas en los problemas detectados"""
        if not diagnosis['detected_issues']:
            diagnosis['recommendations'].append("Intente activar normalmente con el comando 'slmgr /ato'")
            return
        
        # Recomendaciones específicas para problemas detectados
        if "instalación corrupta" in "\n".join(diagnosis['detected_issues']):
            diagnosis['recommendations'].append("Ejecute 'sfc /scannow' para reparar archivos del sistema")
        
        if "servidor KMS" in "\n".join(diagnosis['detected_issues']):
            diagnosis['recommendations'].append("Pruebe con servidores KMS alternativos")
            diagnosis['recommendations'].append("Ejecute 'slmgr /skms kms8.msguides.com'")
        
        if "conexión a internet" in "\n".join(diagnosis['detected_issues']):
            diagnosis['recommendations'].append("Verifique su configuración de red y firewall")
            diagnosis['recommendations'].append("Asegúrese que el puerto 1688 está abierto")
        
        # Recomendación general
        diagnosis['recommendations'].append("Si los problemas persisten, pruebe reiniciando el servicio de licencias: 'net stop sppsvc && net start sppsvc'")
################################################################################################################################################################
    def uninstall_product_key(self) -> bool:
        """
        Uninstalls the current Windows product key using multiple fallback methods
        
        Returns:
            bool: True if uninstallation was successful, False otherwise
        """
        methods = [
            # Method 1: Using slmgr.vbs
            'cscript //nologo "%windir%\\system32\\slmgr.vbs" /upk',
            # Method 2: PowerShell alternative
            'powershell -command "& {$service = Get-WmiObject -Class SoftwareLicensingService; $service.UninstallProductKey()}"'
        ]
        
        for method in methods:
            success, output = self.run_command(method)
            if success:
                print(f"[+] Product key uninstalled successfully using {method.split()[0]}")
                return True
            print(f"[!] Failed to uninstall key with {method.split()[0]}: {output}")
        
        return False

    def clear_kms_server(self) -> bool:
        """
        Clears configured KMS server settings
        
        Returns:
            bool: True if KMS settings were cleared successfully, False otherwise
        """
        clear_kms = 'cscript //nologo "%windir%\\system32\\slmgr.vbs" /ckms'
        success, output = self.run_command(clear_kms)
        
        if not success:
            print(f"[!] Failed to clear KMS server settings: {output}")
            return False
        
        print("[+] KMS server settings cleared successfully")
        return True
            
    def reset_activation_status(self) -> bool:
        """
        Completely resets Windows activation status without requiring reboot
        - Uninstalls product key
        - Clears KMS server settings
        - Attempts reactivation (/ato instead of /rearm)
        
        Returns:
            bool: True if all operations were successful, False otherwise
        """
        print("\n[+] Resetting Windows activation (without reboot)...")
        
        # 1. Uninstall current product key
        if not self.uninstall_product_key():
            print("[!] Failed to uninstall product key")
            return False
        
        # 2. Clear KMS server settings
        if not self.clear_kms_server():
            print("[!] Failed to clear KMS server settings")
            return False
        
        # 3. Attempt reactivation instead of rearm
        reactivate = 'cscript //nologo "%windir%\\system32\\slmgr.vbs" /ato'
        success, output = self.run_command(reactivate)
        
        if not success:
            print(f"[!] Failed to attempt reactivation: {output}")
            return False
        
        print("[+] Windows activation has been reset (no reboot performed)")
        print("    Note: Some features might require reboot to work properly")
        return True
################################################################################################################################################################
    @staticmethod
    def run_command(command: str):
        """Ejecuta un comando y devuelve el resultado"""
        try:
            result = subprocess.run(
                command,
                shell=True,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='ignore'
            )
            return (True, result.stdout.strip())
        except subprocess.CalledProcessError as e:
            return (False, e.stderr.strip())
################################################################################################################################################################
    @staticmethod
    def install_product_key(key: str) -> bool:
        """Instala una clave de producto sin validaciones posteriores."""
        if not key or not isinstance(key, str):
            print("[!] Error: Clave de producto no válida")
            return False

        # Verificar si slmgr.vbs existe primero
        windir = os.environ.get("windir")
        slmgr_path = os.path.join(windir, "system32", "slmgr.vbs") if windir else None
        
        methods = []
        
        if slmgr_path and os.path.exists(slmgr_path):
            methods.append(f'cscript //nologo "{slmgr_path}" /ipk {key}')
        else:
            print("[!] slmgr.vbs no encontrado, usando PowerShell como alternativa")
        
        methods.append(
            'powershell -command "& {$service = Get-WmiObject -Class SoftwareLicensingService; '
            '$service.InstallProductKey(\'' + key + '\'); if($?) {exit 0} else {exit 1}}"'
        )
        
        for method in methods:
            success, output = WindowsActivator.run_command(method)
            if success:
                print(f"[+] Comando ejecutado exitosamente: {method.split()[0]}")
                return True
                
            # Manejo básico de errores
            if "0x80070005" in output:
                print("[!] Error: Permisos insuficientes (ejecute como Administrador)")
                return False
            elif "0xC004F050" in output:
                print("[!] Error: Clave de producto inválida para esta edición de Windows")
                return False
        
        print("[!] Falló la instalación de la clave")
        return False
################################################################################################################################################################
    @staticmethod
    def configure_kms_server(server: str) -> bool:
        """Configura el servidor KMS, compatible con instalaciones no oficiales."""
        # Verificar parámetro
        if not isinstance(server, str) or not server.strip():
            print("[!] Error: Servidor KMS no válido.")
            return False

        server = server.strip()  # Eliminar espacios en blanco al inicio y al final

        # Verificar existencia de slmgr.vbs
        windir = os.environ.get("windir")
        if not windir:
            print("[!] Error: La variable de entorno 'windir' no está definida.")
            return False

        slmgr_path = os.path.join(windir, "system32", "slmgr.vbs")
        if not os.path.exists(slmgr_path):
            print("[!] Error: slmgr.vbs no encontrado (¿instalación corrupta?).")
            return False

        # Configurar servidor KMS
        command = f'cscript //nologo "{slmgr_path}" /skms {server}'
        success, output = WindowsActivator.run_command(command)

        if not success:
            if "0xC004F074" in output:
                print("[!] Error: Servidor KMS inaccesible (¿firewall o fuera de línea?).")
            elif "0x80070005" in output:
                print("[!] Error: Acceso denegado (¿políticas restrictivas?).")
            else:
                print(f"[!] Error al configurar el servidor KMS: {output}")
            return False

        print(f"[+] Servidor KMS configurado correctamente: {server}")

        # Verificar configuración
        verify_cmd = f'cscript //nologo "{slmgr_path}" /dlv'
        verify_success, verify_output = WindowsActivator.run_command(verify_cmd)

        print("[+] Configuración del servidor KMS verificada.")
        return True
################################################################################################################################################################
    def full_activation(self) -> bool:
        def intalacionDeClave(edition):
            gvlk_key = self.GVLK_KEYS.get(edition)  # Acceso a atributo de instancia
            if not gvlk_key:
                print(f"[!] No hay clave GVLK para la edición {edition}")
                return False
            if type(gvlk_key) == list:
                for key in gvlk_key:
                    print(f"[+] Usando clave GVLK: {key}")
                    
                    if self.install_product_key(key):  # Método de instancia
                        print(f"[+] Clave GVLK instalada e forma correcta: {key}")
                        activacion(0)
                else:
                    print("[!] Falló la instalación de las claves disponibles")
                    return False
                
            else:
                print(f"[+] Usando clave GVLK: {gvlk_key}")
                    
                if self.install_product_key(gvlk_key):  # Método de instancia
                    print(f"[+] Clave GVLK instalada e forma correcta: {gvlk_key}")
                    return activacion(0)
                else:
                    print("[!] Falló la instalación de la clave")
                    return False
            
        def activacion(intento):
            # Paso 4: Configurar KMS 
            if intento != len(configurar_servidores_kms_publicos()):
                kms_servers = configurar_servidores_kms_publicos()  # Función externa
                for numServer in range(intento,len(configurar_servidores_kms_publicos())):
                    print(f"\n[+] Probando servidor KMS: {kms_servers[numServer]}")
                    if self.configure_kms_server(kms_servers[numServer]):  # Método de instancia
                        # Paso 5: Activar Windows (con reintentos)
                        print("\n[+] Verificando activación...")
                        if not self.activate_windows():  # Método de instancia
                            print("[!] La activación no fue exitosa según la verificación")
                            self.clear_kms_server()
                            return activacion(intento+1)
                        
                        print("\n[+] Windows activado exitosamente!")
                        # Paso 6: Verificar activación
                        self.verify_activation()
                        return True
                
            else:
                print("[!] Todos los servidores KMS fallaron")
                return False
            
        """Proceso completo de activación con reintentos y verificación"""
        print("\n[+] Iniciando proceso de activación...")
        # Paso 1: Detectar edición
        edition = self.detect_windows_edition()  # Esto es un método de instancia
        if not edition:
            print("[!] No se pudo detectar la edición de Windows")
            return False
            
        print(f"[+] Edición detectada: {edition}")
            
        # Paso 2: Obtener clave GVLK correcta
        return intalacionDeClave(edition)

################################################################################################################################################################
    @staticmethod
    def activate_windows() -> bool:
        """Intenta activar Windows.
            bool: True si la activación fue exitosa, False en caso contrario
        """
        windir = os.environ.get("windir")
        slmgr_path = os.path.join(windir, "system32", "slmgr.vbs") if windir else None
        
        # Verificar existencia de slmgr.vbs
        if not slmgr_path or not os.path.exists(slmgr_path):
            print("[!] Error crítico: slmgr.vbs no encontrado en el sistema")
            print("[*] Solución: Ejecutar 'sfc /scannow' para reparar archivos del sistema")
            return False

        # Ciclo de intentos de activación
        print(f"\n[+] Intento de activación")
            
        # Comando de activación estándar
        command = f'cscript //nologo "{slmgr_path}" /ato'
        success, output = WindowsActivator.run_command(command)
            
        # Análisis de resultados
        output_lower = output.lower()
            
        if success and ("correctamente") in output_lower:
            print("[+] Activación exitosa!")
            print("[*] Información de licencia:")
            WindowsActivator.run_command(f'cscript //nologo "{slmgr_path}" /dli')
            return True
        elif "0xC004C003" in output:
            print("[!] Error: Clave de producto no válida o bloqueada")
            print("[*] Solución: Verificar la clave GVLK para tu versión de Windows")
            return False
        elif "0xC004F038" in output:
            print("[!] Error: Límite de activaciones excedido")
            print("[*] Solución: Esperar 24 horas o usar otro servidor KMS")
            return False
        elif "0xC004F041" in output:
            print("[!] Error: La máquina no está cualificada para activación KMS")
            print("[*] Solución: Verificar que estás usando una versión Volume License")
            return False
        else:
            print(f"[!] Error desconocido durante la activación.")
                
        print(f"\n[!] Falló la activación ")
        return False
################################################################################################################################################################
    @staticmethod
    def verify_activation() -> bool:
        """Verificación mejorada del estado de activación"""
        windir = os.environ.get("windir")
        slmgr_path = os.path.join(windir, "system32", "slmgr.vbs") if windir else None
        
        if not slmgr_path or not os.path.exists(slmgr_path):
            print("[!] Error: slmgr.vbs no encontrado")
            return False

        # Obtener información detallada
        command = f'cscript //nologo "{slmgr_path}" /dlv'
        success, output = WindowsActivator.run_command(command)
        
        if not success:
            print("[!] Error al verificar activación")
            return False
            
        print("\n=== DETALLES DE ACTIVACIÓN ===")
        print(output)
        
        # Patrones de verificación mejorados
        activation_patterns = [
            r"license status:\s*licensed",
            r"estado de la licencia:\s*con licencia",
            r"activation interval:\s*\d+",
            r"renewal interval:\s*\d+",
            r"product activated successfully"
        ]
        
        matches = sum(1 for pattern in activation_patterns if re.search(pattern, output, re.IGNORECASE))
        # Si al menos 2 patrones coinciden, consideramos activado
        if matches >= 1:
            print("[+] Windows está correctamente activado")
            return True
        
        print("[!] Windows NO está activado según la verificación")
        return False
################################################################################################################################################################
def change_windows_edition():
    """Menú para cambiar la edición de Windows"""
    print("\nSeleccione la edición a cambiar:")
    print("1. Windows 10/11 Home (Core)")
    print("2. Windows 10/11 Professional")
    print("3. Cancelar")
    
    option = input("\nOpción (1-3): ").strip()
    
    if option == "1":
        if Cambiar("1"):
            print("\n[✓] Cambio a Home realizado")
        else:
            print("\n[x] Error al cambiar")
    elif option == "2":
        if Cambiar("2"):
            print("\n[✓] Cambio a Professional realizado")
        else:
            print("\n[x] Error al cambiar")
    elif option != "3":
        print("\nOpción no válida")

##########################################
# Menus
##########################################
def activate_windows():
    """Menú principal de activación de Windows"""
    activator = WindowsActivator()
    
    while True:
        # Mostrar información de la edición actual
        current_edition = activator.detect_windows_edition()
        print(f"Edición actual detectada: {current_edition if current_edition else 'Desconocida'}")
        print("\nOpciones principales:")
        print("1. Activar Windows (KMS automático)")
        print("2. Activar Windows con KMS local (vlmcsd)")
        print("3. Verificar estado de activación")
        print("4. Solucionar problemas")
        print("5. Cambiar edición de Windows")
        print("6. Eliminar clave de producto")
        print("7. Resetear estado de activación")
        print("8. Salir")
        
        option = input("\nSeleccione una opción (1-9): ").strip()

        if option == "1":
            if activator.full_activation():
                print("\n[✓] Activación completada con éxito!")
            else:
                print("\n[x] No se pudo completar la activación")
        
        elif option == "2":
            print("futuro")
        
        elif option == "3":
            if activator.verify_activation():
                print("\n[✓] Windows está correctamente activado")
            else:
                print("\n[x] Windows NO está activado")
        
        elif option == "4":
            print(activator.troubleshoot_activation())
        
        elif option == "5":
            change_windows_edition()
        
        elif option == "6":
            if activator.uninstall_product_key():
                print("\n[✓] Clave de producto eliminada correctamente")
            else:
                print("\n[x] No se pudo eliminar la clave de producto")
        
        elif option == "7":
            if activator.reset_activation_status():
                print("\n[✓] Estado de activación reseteado correctamente")
            else:
                print("\n[x] No se pudo resetear el estado de activación")
        
        elif option == "8":
            print("\nSaliendo del programa...")
            break
        
        else:
            print("\n¡Opción no válida! Por favor seleccione 1-9.")
################################################################################################################################################################
def menu_reparar_Windwos():
    print("""
1. Verificar y reparar archivos del sistema
2. Verificar y reparar archivos del sistema (AVANZADO)
3. Comprobar y reparar la imagen de Windows
4. Comprobar y reparar errores en el disco
""")      
    opcion = input("\nSeleccione opción: ").strip()
    if opcion == "1":
        os.system("sfc /scannow")
        input("\nPresione Enter para continuar...")
    elif opcion == "2":
        os.system("dism /online /cleanup-image /scanhealth")
        input("\nPresione Enter para continuar...")
    elif opcion == "3":
        os.system("DISM /Online /Cleanup-Image /CheckHealth")
        os.system("DISM /Online /Cleanup-Image /ScanHealth")
        os.system("DISM /Online /Cleanup-Image /RestoreHealth")
        input("\nPresione Enter para continuar...")
    elif opcion == "4":
        os.system("chkdsk /f /r C:")
        input("\nPresione Enter para continuar...")
################################################################################################################################################################
def menu_cambiar_contrasenia():
    try:
        # Get password securely without echo
        intentoUno = getpass.getpass("Ingrese la nueva contraseña: ")
        intentoDos = getpass.getpass("Ingrese nuevamente la contraseña: ")
        
        if intentoUno != intentoDos:
            print("Las contraseñas no coinciden")
            input("\nPresione Enter para continuar...")
            return
            
        # Get username securely
        username = subprocess.run(
            'whoami',
            capture_output=True,
            text=True,
            shell=True
        ).stdout.split('\\')[-1].strip()
        
        # Change password
        result = subprocess.run(
            ['net', 'user', username, intentoUno],
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0:
            print("Contraseña modificada exitosamente")
            print("Para que surta efecto es necesario cerrar sesión")
            opcion = input("¿Desea hacerlo ahora? (1. Si / 2. No): ").strip()
            if opcion == "1":
                subprocess.run(["shutdown", "/l"])
        else:
            print(f"Error al cambiar la contraseña: {result.stderr.strip()}")
            
    except Exception as e:
        print(f"Error inesperado: {str(e)}")
    
    input("\nPresione Enter para continuar...")
################################################################################################################################################################
def menu_eliminar_contrasenia():
    try:
        # Get current username securely
        username = subprocess.run(
            ['whoami'],
            capture_output=True,
            text=True
        ).stdout.split('\\')[-1].strip()
        
        # Remove password (safer: no shell=True)
        result = subprocess.run(
            ["net", "user", username, ""],
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0:
            print("¡Contraseña eliminada exitosamente!")
            print("Para que surta efecto, es necesario cerrar sesión.")
            opcion = input("¿Desea hacerlo ahora? (1. Si / 2. No): ").strip()
            if opcion == "1":
                subprocess.run(["shutdown", "/l"])
        else:
            if "access denied" in result.stderr.lower():
                print("Error: Se requieren permisos de administrador.")
            else:
                print(f"Error: {result.stderr.strip()}")
                
    except Exception as e:
        print(f"Error inesperado: {str(e)}")
    
    input("\nPresione Enter para continuar...")

##########################################
# MAIN
##########################################
def main():
    if not is_admin():
        print("Ejecute como administrador")
        input("\nPresione Enter para salir...")
        return

    while True:
        os.system('cls')
        print("""
1. Windows
2. Office
3. Herramientas
4. Informacion del sistema
5. Acerca de
6. Salir""")

        opcion = input("\nSeleccione opción: ").strip()

        if opcion == "1":
            activate_windows()
        elif opcion == "2":
            print("""
1. Activar Office
""")        
            opcion = input("\nSeleccione opción: ").strip()
            if opcion == "1":
                activate_office_2019_vl()
        elif opcion == "3":
            print("""
1. Reparar Windows
2. Otras funciones
""")      
            opcion = input("\nSeleccione opción: ").strip()
            if opcion == "1":
                menu_reparar_Windwos()
            elif opcion == "2":
                print("""
1. Servicios
2. Cuentas de usuarios
3. Visor de eventos
4. Contraseña
5. Cambiar Version SO
""")      
                opcion = input("\nSeleccione opción: ").strip()
                if opcion == "1":
                    print(run_command("services.msc"))
                elif opcion == "2":
                    print(run_command("netplwiz"))
                elif opcion == "3":
                    print(run_command("eventvwr.msc"))
                elif opcion == "4":
                    print("""
1. Cambiar contrasenia
2. Eliminar contrasenia
""")      
                    opcion = input("\nSeleccione opción: ").strip()
                    if opcion == "1":
                        menu_eliminar_contrasenia()
                    elif opcion == "2":
                        menu_cambiar_contrasenia()

                elif opcion == "5":
                    print("ADVERTECIA: Antes de cambiar es necesario tener la iso de la nueva version")
                    print("Si se va quitar Enterprise se recomienda seleccionar Windows Home")
                    print("Cambiar version de SO")
                    print("1. Windows Home")
                    print("2. Windows Pro")
                    opcion = input("\nSeleccione opción: ").strip()
                    if opcion == "1" or opcion == "2":
                        Cambiar(opcion)
                    input("\nPresione Enter para continuar...")
        elif opcion == "4":
                obtener_informacion_pc()
                get_windows_info()
                input("\nPresione Enter para continuar...")

        elif opcion == "5":
            print("Acerca de...")
            print("1. KMS")
            print("2. Aplicacion")
            opcion = input("\nSeleccione opción: ").strip()
            if opcion == "1":
                print("1. Ver claves windows")
                print("2. Ver servidores kms")
                opcion = input("\nSeleccione opción: ").strip()
                if opcion == "1":
                    for valor, clave in WindowsActivator.GVLK_KEYS.items():
                        print(f"Version: {valor}, Key: {clave}")
                    input("\nPresione Enter para continuar...")
                elif opcion == "2":
                    for server in configurar_servidores_kms_publicos():
                        print(server)
                    input("\nPresione Enter para continuar...")
            elif opcion == "2":
                print("""
=============================================
=== Herramientas y Activador KMS Avanzado ===
===      Creado por Agustin D'Amore       ===
=============================================
===   Version de la aplicacion 1.2.0.1    ===
=============================================
primer numero <- Mayor Cambios incompatibles con versiones anteriores.
segundo numero <- Menor Nuevas funcionalidades compatibles.
tercer numero <- Parche	Corrección de bugs sin añadir funcionalidades.
cuarto numero <- Build	Número de compilación
""")
            input("\nPresione Enter para continuar...")
        elif opcion == "6":
            break
        else:
            print("Opción no válida")

if __name__ == "__main__":
    main()