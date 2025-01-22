from django.shortcuts import render, redirect
from django.contrib import messages
from django.http import HttpResponse
import pandas as pd
from .classes import ExcelExtract, Screws, Styles

ex = ExcelExtract()

def handle_response(response: dict, request, error_template, data_key='ata'):
    """
    View method for response validation and errors statements

    Args:
        response (dict): Dict response about status, data and errors
        request: Django HTTP Object
        error_template (str): Template to show error
        data_key (str): Dict key to access the data

    Returns:
        object: response data
        HttpResponse: Navigate to error template
    """
    try:
        if response['status']:
            return response[data_key]
        else:
            return render(request, error_template, {'error': response['error']})
    
    except KeyError as err:
        return render(request, 'error.html', {'error': f'KeyError encontrado: {str(err)}'})
        
    except Exception as err:
        return render(request, 'error.html', {'error': f'Erro Inesperado: {str(err)}'})

def upload_files(request):

    """
    View method for upload files in server to manipulation

    Args:
        request (str): Client request to server response

    Returns:
        HttpReponse: Server response
    """

    # Verify if request method is POST and files is not none
    if request.method == 'POST' and request.FILES.getlist('files'):
        files = request.FILES.getlist('files') # Declare files 
        
        # Pipe response
        response = ex.read_all_files(files, 'Pipe')
        pipeDF = handle_response(response, request, 'error.html')
        if isinstance(pipeDF, HttpResponse):
            return pipeDF
        
        # Equipment response
        response = ex.read_all_files(files, 'Piping and Equipment')
        equipDF = handle_response(response, request, 'error.html')
        if isinstance(equipDF, HttpResponse):
            return equipDF
        
        # Concatenando DataFrames
        response = ex.concat(pipeDF, equipDF)
        mainDF = handle_response(response, request, 'error.html')
        if isinstance(mainDF, HttpResponse):
            return mainDF
        
        print(mainDF)

        # try:
        #     df = pd.read_excel(files)
        #     context = []

        #     for _, row in df.iterrows():

        #         new_row = {
        #             'Long Description (Size)': row['Long Description (Size)'],
        #             'Size': row['Size'],
        #             'Spec': row['Spec'],
        #         }

        #         context.append(new_row)

        #     return render(request, 'sucess.html', {'message': 'Excel lido com sucesso!', 'data': context})
        
        # except FileNotFoundError as err:
        #     return render(request, 'upload_files.html', {'error': f'Erro ao encontrar arquivo: {str(err)}'})
        
        # except Exception as err:
        #     return render(request, 'upload_files.html', {'error': f'Erro Inesperado: {str(err)}'})
    else:
        return render(request, 'upload_files.html')