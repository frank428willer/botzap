import os
from flask import Flask, request, jsonify
from twilio.twiml.messaging_response import MessagingResponse
import openpyxl

app = Flask(__name__)

#carregar planilha 
workbook = openpyxl.load_workbook('planilha.xlsx') 

if 'planilha.xlsx' in os.listdir() :
    sheet = workbook.active

@app.route('/bot', methods=['POST'])
def bot():
    incoming_msg = request.values.get('Body', '').strip()
    resp = MessagingResponse()
    msg = resp.message()
    
    if incoming_msg:
        # Supondo que a mensagem tenha o formato "Nome, Idade, Cidade"
        data = incoming_msg.split(',')
        if len(data) == 3:
            nome, idade, cidade = data
            # Adicionar dados à planilha
            sheet.append([nome.strip(), idade.strip(), cidade.strip()])
            workbook.save('planilha.xlsx')
            response_msg = f"Dados salvos: Nome={nome}, Idade={idade}, Cidade={cidade}"
        else:
            response_msg = "Formato inválido. Use: Nome, Idade, Cidade"
    else:
        response_msg = "Mensagem vazia."

    msg.body(response_msg)
    return str(resp)

if __name__ == '__main__':
    app.run(debug=True)
