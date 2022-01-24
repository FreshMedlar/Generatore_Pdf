birth = "1984/12/24"
year = birth.partition('/')[0]

month = birth.partition('/')[2].partition('/')[0]

day = birth.partition('/')[2].partition('/')[2].partition('/')[0]

print(f'{day}/{month}/{year}')
