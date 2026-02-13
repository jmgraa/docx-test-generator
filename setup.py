import random
from pathlib import Path
from shutil import rmtree

from builder import (
    add_question_on_document,
    create_document,
    save_data_file,
    save_document,
)
from parsing import parse_exam_file
from shuffling import shuffle_answers


def mix_tests(input_params):
    docx_dir = Path(input_params.output_dir) / "Testy"
    dat_dir = Path(input_params.output_dir) / "KluczeDAT"
    txt_dir = Path(input_params.output_dir) / "KluczeTXT"
    images_dir = Path(input_params.output_dir) / "Images"
    docx_dir.mkdir()
    dat_dir.mkdir()
    txt_dir.mkdir()
    images_dir.mkdir()

    questions = parse_exam_file(input_params.file_path, images_dir)
    number_of_questions = len(questions)

    for i in range(1, input_params.copy_number + 1):
        yield i
        shuffled_questions = random.sample(questions, number_of_questions)
        dat_data = []
        txt_data = []
        docx = create_document(
            f"{input_params.title}, {input_params.date}{" " * 5}Test {str(i).zfill(3)}",
            input_params.intro,
        )

        for q_num, q in enumerate(shuffled_questions):
            shuffled_answers, order = shuffle_answers(q)
            docx = add_question_on_document(docx, q_num + 1, q, shuffled_answers)
            dat_data.append(f"{q.id} {" ".join(str(x + 1) for x in order)}")
            txt_data.append(
                f"{q_num + 1} {" ".join(str(x + 1) if not x else '0' for x in order)} {q.weight}"
            )

        save_document(docx, i, docx_dir)
        save_data_file("k", "dat", dat_data, i, dat_dir)
        save_data_file("klucz", "txt", txt_data, i, txt_dir)

    rmtree(images_dir)
