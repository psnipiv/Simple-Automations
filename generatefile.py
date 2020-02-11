import pandas as pd


# To modify the card details, start and ending card no and number of cards in a file
card_prefix = '00002'
card_suffix_length = 14
membership_type = 'Platinum'
membership_tier = 'Member'
startno = 1
endno = 10000
no_records = 5000

# Function to generate Excel file for preload cards
def generate_file(a,b):
    filename = str(a) + "-" + str(b-1) + ".xlsx"
    pvlist = [str(x).zfill(card_suffix_length) for x in range(a,b,1)]
    df = pd.DataFrame({'CardSuffix':pvlist})
    df['Prefix'] = card_prefix
    df['Card Number'] = df['Prefix'] + df['CardSuffix']
    df['Membership Type'] = membership_type
    df['Membership Tier'] = membership_tier
    df['Cardno_Length'] =df['Card Number'].apply(len)
    finaldf = df[['Prefix','CardSuffix','Cardno_Length','Card Number','Membership Type','Membership Tier']]
    finaldf.to_excel(filename)
    return "File " + filename + " is created."

def main():
    lst = list(range(startno,endno+1,no_records))
    print(lst)
    lst_range1_length = len(list(range(startno,endno,1)))
    if lst_range1_length > no_records:
        for a, b in zip(lst[:-1], lst[1:]):
            result = generate_file(a,b)
            print(result)
            print("\n")
        rangeendno = lst[-1]
        if lst[-1] == endno:
            print("File generation is completed")
        else:
            result = generate_file(rangeendno,endno+1)
            print(result)
            print("\n")
            print("File generation is completed")
    else:
        print("No of cards in a file should be greater than the total number of cards given the range")


if __name__== "__main__":
    main()