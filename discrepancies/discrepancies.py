from printing import print_colour, GREEN


def print_discrepancies(discrepancies):
    if discrepancies:

        print("Mismatches found:")
        for d in discrepancies:
            print(d)
    else:
        print_colour(GREEN, "No mismatches found.")
