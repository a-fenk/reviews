class Config:
    LIMIT_MASTERS = None

    MASTERS_SHEET = 'Masters_URL'
    REVIEWS_SHEET = 'All_reviews'
    SC_SHEET = 'SC'
    TAGS_SHEET = 'Шаблон тегов'

    TAG_WORD_COLUMN = 'Слова тегов'
    TAG_NAME_COLUMN = 'Name_tag'
    TAG_LEVEL_COLUMN = 'уровень'

    REVIEWS_COLUMNS = [
        'Отзыв',
        'Masters_URL',
        'ID section',
        'ID container',
        'Name section',
        '№ заказа',
        'Гео район',
        'Гео метро',
        'Corrected',
        'Кол-во отзывов',
        'Кол-во отзывов Corrected - TRUE',
    ]
    REVIEWS_SEARCH_RANGE = {
        'from': 'B',
        'to': 'B',
    }

    SC_COLUMNS = [
        'id container',
        'Address',
        'H1-1',
    ]
    SC_SEARCH_RANGE = {
        'from': 'EF',
        'to': 'FE',
    }

    RESULT_COLUMNS = [
        'Отзыв',
        'Masters_URL',
        'ID section',
        'ID container',
        'Name section',
        '№ заказа',
        'Гео район',
        'Гео метро',
        'Corrected',
        'Кол-во отзывов',
        'Кол-во отзывов Corrected - TRUE',
        'Address',
        'H1-1',
        'name_tag'
    ]

    RESULT_FILE_NAME = 'result.xlsx'
    SOURCE_FILE_NAME = 'source.xlsx'

