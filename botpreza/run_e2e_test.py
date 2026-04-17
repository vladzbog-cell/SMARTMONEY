import asyncio
import os
import time

from ai_service import analyze_and_create_structure
from presentation_builder import build_pptx_from_json


TEST_TEXT = '''
Нейросети в образовании помогают персонализировать обучение, автоматизировать проверку заданий и анализировать прогресс студентов.

Основные преимущества:
- адаптивные траектории обучения;
- быстрое выявление пробелов в знаниях;
- снижение нагрузки на преподавателей;
- интерактивные форматы объяснения сложных тем.

Основные риски:
- зависимость от качества исходных данных;
- возможные ошибки в рекомендациях;
- вопросы приватности и хранения данных;
- необходимость контроля со стороны преподавателя.

Сделай деловую презентацию для руководства вуза.
'''


async def main() -> None:
    started = time.time()
    print('=== ANALYZE START ===')
    json_result = await analyze_and_create_structure(TEST_TEXT, 'academic')
    print('=== ANALYZE DONE ===')
    print(json_result[:2000])

    output_path = '/opt/prezbot/temp/e2e_test_3slides.pptx'
    if os.path.exists(output_path):
        os.remove(output_path)

    print('=== BUILD START ===')
    result_path = await build_pptx_from_json(json_result, output_path)
    print('=== BUILD DONE ===')
    print(f'RESULT_PATH={result_path}')
    print(f'EXISTS={os.path.exists(result_path)}')
    if os.path.exists(result_path):
        print(f'SIZE={os.path.getsize(result_path)}')
    print(f'DURATION_SEC={round(time.time() - started, 2)}')


if __name__ == '__main__':
    asyncio.run(main())