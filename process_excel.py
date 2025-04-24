
def Main(fileLoad):
  from utils import load_file, copy_sheet, index_column_address, insert_column, organize_worksheet, remove_duplicate, remove_apartment, save_workbook, path_project, edit_length, check_address, formatDateFile, buscaCEP, verifyStreetNameByZipCode, correct_address, remove_columns

  # Ler o arquivo
  _, worksheet = load_file(fileLoad)
  
  # Faz uma copia do arquivo
  data = copy_sheet(worksheet)
  
  # Remove as colunas: AT ID; Sequence; Stop
  remove_columns(worksheet, data, ['AT ID', 'Sequence', 'Stop'])


  # Faz uma nova copia do arquivo
  data = copy_sheet(worksheet)
  

  # Descobre a coluna que fica o endereço, descrição
  indexColumnAddress = index_column_address(data)
  indexColumnDescription = indexColumnAddress + 1
  indexColumnAmount = indexColumnDescription + 1


  # Divide o address em "rua, numero" e referencia
  organize_worksheet(data, indexColumnDescription, indexColumnAddress)

  # Insere a coluna 'quantidade' no final da linha
  insert_column(data, indexColumnAmount, 'QUANTIDADE', 1)

  # Corrige o nome dos endereços === Fase de testes
  # correct_address(data, indexColumnAddress)
  
  # corrige os endereços errados que começam com "r. ", "R. " por "Rua"
  check_address(data, indexColumnAddress)
  
  data = remove_duplicate(data, indexColumnAddress, indexColumnAmount)
  
  # data = fetchAddress(data, indexColumnAddress)

  # buscaCEP(data)

  # remove os endereços do por do sol e add uma parada somente com a quantidade
  remove_apartment(data, indexColumnAddress, indexColumnAmount)

  edit_length(data, indexColumnAmount)

  # pasta local / data atual
  name_file_posfixo = path_project()
  name_file = f'{name_file_posfixo}/{formatDateFile()}'

  save_workbook(data, name_file)

  return name_file
