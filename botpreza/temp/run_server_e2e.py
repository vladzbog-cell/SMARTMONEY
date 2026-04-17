import asyncio
from pathlib import Path
from ai_service import analyze_and_create_structure
from presentation_builder import build_pptx_from_json

TEXT = """
Суть нашего NotebookLM для презентаций строится на трех опорах. Первая — абсолютная всеядность: система переваривает DOCX, PDF, голосовые сообщения, фотографии конспектов, скриншоты графиков и ссылки на статьи. Вторая — глубокое понимание контекста: для комиссии нужен data-driven подход с матрицами и сравнением сценариев, а для питча стартапа — более метафоричная и минималистичная подача. Третья — премиальная гибридная верстка: ИИ создает сильный визуальный фон, а программная сборка накладывает типографику, карточки, схемы и аккуратные контейнеры. В результате презентация должна выглядеть как дорогой strategy report, а не как стандартный шаблон PowerPoint.
"""

async def main():
    structure = await analyze_and_create_structure(TEXT, "academic")
    print("STRUCTURE_START")
    print(structure)
    print("STRUCTURE_END")
    output = await build_pptx_from_json(structure, output_filename="server_live_e2e_3slides.pptx")
    path = Path(output)
    print("OUTPUT", output)
    print("EXISTS", path.exists())
    if path.exists():
        print("SIZE", path.stat().st_size)

asyncio.run(main())
