try:
    from aiohttp import web
except ImportError:
    raise SystemExit("Install with server extra to use this command.")
import importlib.metadata
import logging
import os.path
import tempfile

from docx import Document

from docxcompose.composer import Composer


CHUNK_SIZE = 65536
logger = logging.getLogger("docxcompose")
version = importlib.metadata.version("docxcompose")


async def compose(request):

    documents = []

    if not request.content_type == "multipart/form-data":
        logger.info(
            "Bad request. Received content type %s instead of multipart/form-data.",
            request.content_type,
        )
        return web.Response(status=400, text="Multipart request required")

    reader = await request.multipart()

    with tempfile.TemporaryDirectory() as temp_dir:
        while True:
            part = await reader.next()

            if part is None:
                break

            if part.filename is None:
                continue

            documents.append(await save_part_to_file(part, temp_dir))

        if not documents:
            return web.Response(status=400, text="No documents provided")

        composed_filename = os.path.join(temp_dir, "composed.docx")

        try:
            composer = Composer(Document(documents.pop(0)))
            for document in documents:
                composer.append(Document(document))
            composer.save(composed_filename)
        except Exception:
            logger.exception("Failed composing documents.")
            return web.Response(status=500, text="Failed composing documents")

        return await stream_file(
            request,
            composed_filename,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


async def save_part_to_file(part, directory):
    filename = os.path.join(directory, f"{part.name}_{part.filename}")
    with open(filename, "wb") as file_:
        while True:
            chunk = await part.read_chunk(CHUNK_SIZE)
            if not chunk:
                break
            file_.write(chunk)
    return filename


async def stream_file(request, filename, content_type):
    response = web.StreamResponse(
        status=200,
        reason="OK",
        headers={
            "Content-Type": content_type,
            "Content-Disposition": f'attachment; filename="{os.path.basename(filename)}"',  # noqa
        },
    )
    await response.prepare(request)

    with open(filename, "rb") as outfile:
        while True:
            data = outfile.read(CHUNK_SIZE)
            if not data:
                break
            await response.write(data)

    await response.write_eof()
    return response


async def healthcheck(request):
    return web.Response(status=200, text="OK")


def create_app():
    app = web.Application()
    app.add_routes([web.post("/", compose)])
    app.add_routes([web.get("/healthcheck", healthcheck)])
    return app


def main():
    print(f"docxcompose {version}")
    logging.basicConfig(
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        level=logging.INFO,
    )
    web.run_app(create_app())


if __name__ == "__main__":
    main()
