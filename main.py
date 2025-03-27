from xlsx_parser import ParserXLSX
from os import listdir
from re import compile, findall, sub


def get_new_file_name_pattern() -> str:
    pattern_name = ''
    while len(findall(compile('\\{(.+)?\\}'), pattern_name)) == 0:
        pattern_name = input('''please enter file name template in format "filename{year}". 
            The year (int) will be inserted in {}
            Only .xlsx files!\t''')
    pattern_name = sub(compile('\\{(.+)?\\}'), '{year}', pattern_name)
    pattern_name = pattern_name.split('.')[0]
    return pattern_name + '.xlsx'

def get_available_years(file_name_pattern: str) -> list:
    list_ = []
    file_name_pattern = sub(compile('{year}'), r'(\\d{4})', file_name_pattern)
    for i in listdir('files_to_parse'):
        var = findall(compile(file_name_pattern), i)
        if var:
            list_.append(var[0])
    return list_


if __name__ == '__main__':
    file_name_pattern = 'population_belarus_{year}.xlsx'
    var = input(f'file name template is "{file_name_pattern}"? ' + 'The year (int) will be inserted in {}\ny/n\t')
    while var != 'y':
        if var != 'n':
            var = input('please enter y or n\t')
        else:
            file_name_pattern = get_new_file_name_pattern()
            var = input(f'file name template is "{file_name_pattern}"? ' + 'The year (int) will be inserted in {}\ny/n\t'
                        )
    years = get_available_years(file_name_pattern)
    print(f'The following years are available: {', '.join(years)}')
    year = input('Please enter year: ').replace(' ', '')
    while year not in years:
        year = input(f'The file with the year {year} is not in the folder. Please enter year: ').replace(' ', '')
    df = ParserXLSX(year=int(year), file_name_pattern=file_name_pattern.replace('{year}', '{}')).parse()
    new_file_name = f'ready_population_belarus_{year}.xlsx'
    df.to_excel(new_file_name, index=False)
    print(f'Please see the result in the "{new_file_name}" file')



