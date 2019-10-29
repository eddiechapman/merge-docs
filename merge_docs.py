"""
merge_docs.py

Concatenate the mission and skills documents for each degree.

Create a 'description' document is created for degrees that have a valid
skills and valid mission document. The resulting document combines the two.

Incomplete, invalid, course, and unknown documents are copied to new
directories for inspection.

The contents of the input directory are not disturbed.

                        [ID]-[Category].docx

    - [ID] a three digit number (zero-padded) representing a degree
    - [Category] one of "Skills", "Courses", or "Mission"

Output directories:

    - descriptions: new description documents
    - incomplete: mission and skills documents that could not be paired
    - courses: courses documents
    - error: files that had no text or otherwise caused an error
    - unknown: files from the input directory that could not be sorted

This script is part of the PERVADE project for the data science
curriculum sub-project.

"""
import argparse
import logging
import pathlib
import re
import docx


def main(args):
    logging.basicConfig(
        level=args.log_level or logging.INFO,
        filename='app.log',
        format='%(levelname)s:&(asctime)s:%(message)s'
    )

    INPUT = pathlib.Path(args.input)
    OUTPUT = pathlib.Path(args.output)

    if not OUTPUT.exists() or OUTPUT.is_dir():
        OUTPUT.mkdir()

    DESCRIPTIONS = OUTPUT / 'descriptions'
    INCOMPLETE = OUTPUT / 'incomplete'
    COURSES = OUTPUT / 'courses'
    ERROR = OUTPUT / 'error'
    UNKNOWN = OUTPUT / 'unknown'

    DESCRIPTIONS.mkdir()
    INCOMPLETE.mkdir()
    COURSES.mkdir()
    ERROR.mkdir()
    UNKNOWN.mkdir()

    degree_docs = dict()
    pattern = r'(?P<id>\d{3})-(?P<category>Skills|Courses|Mission).docx'

    for filename in INPUT.glob('*.docx'):
        logging.debug(f'Sorting file: {filename}')
        m = re.find(pattern, filename)
        if m:
            logging.debug(f'{filename} matched: {m.groupdict()}')
            degree_docs[m.group('id')][m.group('category')] = filename

    for degree, docs in documents:
        skills = docs.get('Skills')
        mission = docs.get('Mission')
        courses = docs.get('Courses')

        if courses:
            course_path = INPUT / courses
            logging.debug(f'Moving {courses} to {COURSES}')
            course_path.rename(COURSES / courses)

        if mission and skills:
            m_doc = docx.Document(mission)
            s_doc = docx.Document(skills)
            m_txt = '\n'.join([p.text for p in m_doc.paragraphs])
            s_txt = '\n'.join([p.text for p in s_doc.paragraphs])

            if m_txt and s_txt:
                path = DESCRIPTIONS / f'{degree}-Description.docx'
                description = docx.Document(path)
                description.add_paragraph(m_txt)
                description.add_paragraph(s_txt)
                description.save()
                logging.debug(f'Description for #{degree} written to {path}')

            else:
                if not m_txt:
                    logging.error(f'Bad document: {mission}')
                    mission_path = INPUT / mission
                    mission_path.rename(ERROR / mission)
                if not s_txt:
                    logging.error(f'Bad document {skills}')
                    skills_path = INPUT / skills
                    skills_path.rename(ERROR / skills)

        else:
            if mission:
                mission_path = INPUT / mission
                mission_path.rename(INCOMPLETE / mission)
                logging.debug(f'Moving incomplete: {mission_path}')
            elif skills:
                skills_path = INPUT / skills
                skills_path.rename(INCOMPLETE / skills)
                logging.debug(f'Moving incomplete: {skills_path}')

    # Separate unknown files from input directory
    known = set()
    for docs in documents.values():
        known.update(docs.values())

    for doc in INPUT.iterdir():
        if not doc.name in known:
            doc.rename(UNKNOWN / doc.name)
            logging.warn(f'Moving unknown file: {doc}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        '-v', '--verbose',
        help='Verbose (debug) logging level.',
        const=logging.DEBUG,
        dest='log_level',
        nargs='?',
    )
    group.add_argument(
        '-q', '--quiet',
        help='Silent mode, only log warnings and errors.',
        const=logging.WARN,
        dest='log_level',
        nargs='?',
    )
    parser.add_argument(
        '-i', '--input',
        help=(f'Path to directory of .docx files in format: '
              f'###-[Skills|Mission|Courses].docx'),
        metavar='input',
        type=str
    )
    parser.add_argument(
        '-o', '--output',
        help=(f'Path to the directory where the resulting files '
              f'will be written.'),
        metavar='output',
        default='./output',
        type=str
    )
    args = parser.parse_args()
    main(args)

