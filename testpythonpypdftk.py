import pypdftk
import os.path
import os
import subprocess
import re
import msvcrt

# variables globales
direc_archivo = 'formulario.pdf'
final_pdf = 'formulario_rellenado.pdf' 

#lista para insetar en pdf
datos_cabeza_pie = {
    "CONTRATISTARow1": "Jose",
"CONTRATORow1": "Mauricio",
"PROCEDIMIENTORow1": "inspección",
"DEL": "45181",
"AL": "45242",
"CLAVERow1": "444DRS",
"CONCEPTORow1": "edifico",
"CONSECUTIVORow1": "342",
"NUMERO DE BACHERow1": "343",
"CONSECUTIVORow2": "344",
"NUMERO DE BACHERow2": "345",
"CONSECUTIVORow3": "346",
"NUMERO DE BACHERow3": "347",
"CONSECUTIVORow4": "348",
"NUMERO DE BACHERow4": "349",
"CONSECUTIVORow5": "350",
"NUMERO DE BACHERow5": "351",
"CONSECUTIVORow6": "352",
"NUMERO DE BACHERow6": "353",
"CONSECUTIVORow7": "354",
"NUMERO DE BACHERow7": "355",
"CONSECUTIVORow8": "356",
"NUMERO DE BACHERow8": "357",
"CONSECUTIVORow9": "358",
"NUMERO DE BACHERow9": "359",
"CONSECUTIVORow10": "360",
"NUMERO DE BACHERow10": "361",
"CONSECUTIVORow11": "362",
"NUMERO DE BACHERow11": "363",
"CONSECUTIVORow12": "364",
"NUMERO DE BACHERow12": "365",
"CONSECUTIVORow13": "366",
"NUMERO DE BACHERow13": "367",
"CONSECUTIVORow14": "368",
"NUMERO DE BACHERow14": "369",
"CONSECUTIVORow15": "370",
"NUMERO DE BACHERow15": "371",
"CONSECUTIVORow16": "372",
"NUMERO DE BACHERow16": "373",
"CONSECUTIVORow17": "374",
"NUMERO DE BACHERow17": "375",
"CONSECUTIVORow18": "376",
"NUMERO DE BACHERow18": "377",
"CONSECUTIVORow19": "378",
"NUMERO DE BACHERow19": "379",
"CONSECUTIVORow20": "380",
"NUMERO DE BACHERow20": "381",
"CONSECUTIVORow21": "382",
"NUMERO DE BACHERow21": "383",
"CONSECUTIVORow22": "384",
"NUMERO DE BACHERow22": "385",
"CONSECUTIVORow23": "386",
"NUMERO DE BACHERow23": "387",
"CONSECUTIVORow24": "388",
"NUMERO DE BACHERow24": "389",
"CONSECUTIVORow25": "390",
"NUMERO DE BACHERow25": "391",
"CONSECUTIVORow26": "392",
"NUMERO DE BACHERow26": "393",
"CONSECUTIVORow27": "394",
"NUMERO DE BACHERow27": "395",
"CONSECUTIVORow28": "396",
"NUMERO DE BACHERow28": "397",
"CONSECUTIVORow29": "398",
"NUMERO DE BACHERow29": "399",
"CONSECUTIVORow30": "400",
"NUMERO DE BACHERow30": "401",
"CONSECUTIVORow31": "402",
"NUMERO DE BACHERow31": "403",
"CONSECUTIVORow32": "404",
"NUMERO DE BACHERow32": "405",
"CONSECUTIVORow33": "406",
"NUMERO DE BACHERow33": "407",
"CONSECUTIVORow34": "408",
"NUMERO DE BACHERow34": "409",
"CONSECUTIVORow35": "410",
"NUMERO DE BACHERow35": "411",
"CLAVE DE BACHERow1": "412",
"ANCHO MRow1": "413",
"LARGO M2Row1": "414",
"AREA M2Row1": "415",
"ANCHO MRow2": "416",
"LARGO M2Row2": "417",
"AREA M2Row2": "418",
"ANCHO MRow3": "419",
"LARGO M2Row3": "420",
"AREA M2Row3": "421",
"ANCHO MRow4": "422",
"LARGO M2Row4": "423",
"AREA M2Row4": "424",
"ANCHO MRow5": "425",
"LARGO M2Row5": "426",
"AREA M2Row5": "427",
"ANCHO MRow6": "428",
"LARGO M2Row6": "429",
"AREA M2Row6": "430",
"ANCHO MRow7": "431",
"LARGO M2Row7": "432",
"AREA M2Row7": "433",
"ANCHO MRow8": "434",
"LARGO M2Row8": "435",
"AREA M2Row8": "436",
"ANCHO MRow9": "437",
"LARGO M2Row9": "438",
"AREA M2Row9": "439",
"ANCHO MRow10": "440",
"LARGO M2Row10": "441",
"AREA M2Row10": "442",
"ANCHO MRow11": "443",
"LARGO M2Row11": "444",
"AREA M2Row11": "445",
"ANCHO MRow12": "446",
"LARGO M2Row12": "447",
"AREA M2Row12": "448",
"ANCHO MRow13": "449",
"LARGO M2Row13": "450",
"AREA M2Row13": "451",
"ANCHO MRow14": "452",
"LARGO M2Row14": "453",
"AREA M2Row14": "454",
"ANCHO MRow15": "455",
"LARGO M2Row15": "456",
"AREA M2Row15": "457",
"ANCHO MRow16": "458",
"LARGO M2Row16": "459",
"AREA M2Row16": "460",
"ANCHO MRow17": "461",
"LARGO M2Row17": "462",
"AREA M2Row17": "463",
"ANCHO MRow18": "464",
"LARGO M2Row18": "465",
"AREA M2Row18": "466",
"ANCHO MRow19": "467",
"LARGO M2Row19": "468",
"AREA M2Row19": "469",
"ANCHO MRow20": "470",
"LARGO M2Row20": "471",
"AREA M2Row20": "472",
"ANCHO MRow21": "473",
"LARGO M2Row21": "474",
"AREA M2Row21": "475",
"ANCHO MRow22": "476",
"LARGO M2Row22": "477",
"AREA M2Row22": "478",
"ANCHO MRow23": "479",
"LARGO M2Row23": "480",
"AREA M2Row23": "481",
"ANCHO MRow24": "482",
"LARGO M2Row24": "483",
"AREA M2Row24": "484",
"ANCHO MRow25": "485",
"LARGO M2Row25": "486",
"AREA M2Row25": "487",
"ANCHO MRow26": "488",
"LARGO M2Row26": "489",
"AREA M2Row26": "490",
"ANCHO MRow27": "491",
"LARGO M2Row27": "492",
"AREA M2Row27": "493",
"ANCHO MRow28": "494",
"LARGO M2Row28": "495",
"AREA M2Row28": "496",
"ANCHO MRow29": "497",
"LARGO M2Row29": "498",
"AREA M2Row29": "499",
"ANCHO MRow30": "500",
"LARGO M2Row30": "501",
"AREA M2Row30": "502",
"ANCHO MRow31": "503",
"LARGO M2Row31": "504",
"AREA M2Row31": "505",
"ANCHO MRow32": "506",
"LARGO M2Row32": "507",
"AREA M2Row32": "508",
"ANCHO MRow33": "509",
"LARGO M2Row33": "510",
"AREA M2Row33": "511",
"ANCHO MRow34": "512",
"LARGO M2Row34": "513",
"AREA M2Row34": "514",
"ANCHO MRow35": "515",
"LARGO M2Row35": "516",
"AREA M2Row35": "517",
"M2 ACOMULADOS": "45534",
"M2 TOTAL DE ESTA HOJA": "65554",
"DESCRIPCI&#211;N": "sona despejada",
"cargo_aprobo": "personal",
"nombre_aprobo": "johan",
"nombre_superviso": "luis",
"cargo_superviso": "administrador",
"nombre_elaboro": "martin",
"cargo_elaboro": "movilidad",
"areaM": "345",
"areaM_": "6554",
    }


#funcion para checar si existe el archivo formulario.pdf
def archivo_existe(direc_archivo):
    return os.path.exists(direc_archivo)


#funcion para extraer los campos del pdf
def get_field_names(direc_archivo):
    try:
        pdftk_comando = ["pdftk", direc_archivo, "dump_data_fields"]
        resultado = subprocess.run(pdftk_comando, stdout=subprocess.PIPE, text=True, check=True)
        
        # se busca que contenga el campo FielName y de ser asi, se captura el dato en la lista
        nombre_datos = resultado.stdout
        nombres_de_campos = re.findall(r"FieldName: (.+)", nombre_datos)

        if nombres_de_campos:
            print("IMPRIMIR CAMPOS DEL PDF \n")
            for extracto_nombres_pdf in nombres_de_campos:
                print(extracto_nombres_pdf)

    except subprocess.CalledProcessError as e:
        print("Error : No se encuentra el archivo =",direc_archivo)
        print(e)
        return None


#funcion para insertar los datos en el pdf y generar nuevo pdf
def generar_pdf(direc_archivo,datos_cabeza_pie,final_pdf):
    try:
        #uso de metodo fill_form para insertar los datos
        generador_pdf = pypdftk.fill_form(direc_archivo, datos_cabeza_pie, final_pdf)
        subprocess.Popen(generador_pdf,shell=True)

    except subprocess.CalledProcessError as b:
        print("Error : No se realizo la insersión de datos correctamente en el archivo=",final_pdf)
        print(b)
        return None


#inicializar script
if __name__ == "__main__":
    #validar si existe archivo formulario.pdf
    if archivo_existe(direc_archivo):
        
        #extraer campos del pdf e imprimirlos en consola
        get_field_names(direc_archivo)

        #insertar datos al pdf y crear pdf
        generar_pdf(direc_archivo,datos_cabeza_pie,final_pdf)

    else:
        print(f"El archivo {direc_archivo} no se encuentra o no esta en la misma ruta que el (script o exe) para ejecutar el programa.")
        
msvcrt.getch()