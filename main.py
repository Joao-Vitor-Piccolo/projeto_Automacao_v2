import asyncio
import utils
import schedule
import time

queue_make_pptx = asyncio.Queue()
queue_convert_pdf = asyncio.Queue()
queue_convert_image = asyncio.Queue()
queue_send_email = asyncio.Queue()


async def process_powerpoint():
    while True:
        data = await queue_make_pptx.get()
        print('Item pego na fila de powerpoint: ', data)
        to_pdf = await utils.make_slide()
        await queue_convert_pdf.put(to_pdf)
        queue_make_pptx.task_done()


async def convert_to_pdf():
    while True:
        file = await queue_convert_pdf.get()
        print('Item pego na fila de PDF: ', file)
        await utils.convert_slide(file)
        await queue_convert_image.put(file.replace("pptx", "pdf"))
        queue_convert_pdf.task_done()


async def convert_to_img():
    while True:
        data = await queue_convert_image.get()
        print('Item pego na fila de imagem: ', data)
        data = await utils.convert_pdf(data)
        await queue_send_email.put(data)
        queue_convert_image.task_done()


async def async_send_email():
    while True:
        file = await queue_send_email.get()
        print('Item pego na fila de email: ', file)
        await utils.send_email(file)
        queue_send_email.task_done()


async def main():
    lista = utils.lista_copy
    asyncio.create_task(process_powerpoint())
    asyncio.create_task(convert_to_pdf())
    asyncio.create_task(convert_to_img())
    asyncio.create_task(async_send_email())
    for x in range(len(lista)):
        await queue_make_pptx.put(x)


schedule.every().day.at(utils.get_time()).do(asyncio.run(main()))

while True:
    schedule.run_pending()
    time.sleep(1)
    if not schedule.get_jobs():
        print("Pronto!")
        break
