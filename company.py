class Company:
    official_name = ''
    trading_name = ''
    bussiness_code = ''
    date_of_license = ''
    start_working_date = ''
    status = ''
    address = ''
    phone = ''
    email = ''
    director = ''
    director_phone = ''
    accountant = ''
    accountant_phone = ''
    business_lines = ''

    def __repr__(self):
        return str(self.__dict__)
