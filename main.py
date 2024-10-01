import asyncio
import utils

# Cria um objeto Queue, ou fila, para cada processo
queue_make_pptx = asyncio.Queue()
queue_convert_pdf = asyncio.Queue()
queue_convert_image = asyncio.Queue()
queue_send_email = asyncio.Queue()


# Essa funções começam os processos de criação de slide e os colocam nas filas criadas

async def async_process_powerpoint():
    # while True vai rodar e vai esperar algum dado nas filas, por isso não precisa de sleep
    while True:
        data = await queue_make_pptx.get()  # Espera o primeiro dado estar disponivel
        print('Item pego na fila de powerpoint: ', data)
        to_pdf = await utils.make_slide()  # esperar até que o slide esteja pronto e
        # vai armazenar o nome dele na variavel
        await queue_convert_pdf.put(to_pdf)  # Vai colocar o "path" do arquivo pptx, na fila de conversão para PDF
        queue_make_pptx.task_done()  # Vai terminar essa task para uma proxima


async def async_convert_to_pdf():
    while True:
        # Pega aqui o "path" passado na finalização da função anterior, porém, já formatado
        file = await queue_convert_pdf.get()
        print('Item pego na fila de PDF: ', file)
        await utils.convert_to_pdf(file)  # Converte de pptx, para PDF
        await queue_convert_image.put(file.replace("pptx", "pdf"))  # Muda a variavel de pptx, para PDF.
        # pois diferente das outras o convert_to_pdf(), converte todos os arquivos .pptx, não um de cada
        # vez então o .replace(), simula isso, enviado o path "convertido", para as proximas filas.
        queue_convert_pdf.task_done()  # Encerra a fila para uma proxima


async def async_convert_to_img():
    while True:
        data = await queue_convert_image.get()  # Pega o path que foi passado pela função assincrona "convert_to_pdf"
        print('Item pego na fila de imagem: ', data)
        data = await utils.convert_to_img(data)  # Converte em imagem, (A variavel vai receber o path novo, ".jpg")
        await queue_send_email.put(data)  # Vai enviar o path para a fila de enviar emails
        queue_convert_image.task_done()  # Encerra a fila para um proximo dado


async def async_send_email():
    while True:
        file = await queue_send_email.get()  # Pega o path da imagem
        print('Item pego na fila de email: ', file)
        await utils.send_email(file)  # Coloca o path na função e a chama
        queue_send_email.task_done()  # Termina a fila para a proxima.


async def main():
    # Cria as "tasks" das funções assincronas, as chamando-as e
    # também fazendo assim a parte assincrona do código funcionar.
    asyncio.create_task(async_process_powerpoint()),
    asyncio.create_task(async_convert_to_pdf()),
    asyncio.create_task(async_convert_to_img()),
    asyncio.create_task(async_send_email())

    # Chama, pelo .utils, uma lista que contem  a quantidade de vezes que o código precisa funcionar
    for x in range(len(utils.lista)):
        await queue_make_pptx.put(x)  # Para cada item na lista, coloque um item na fila de fazer slides

if __name__ == '__main__':
    asyncio.run(main())
