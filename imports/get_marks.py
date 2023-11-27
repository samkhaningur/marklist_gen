def get_mark(name,subject,maxmarks):
    inp=float(input("Enter "+ name +"'s " + subject + " mark: "))
    if inp.isdigit():
        return inp 
    else:
        print("Enter a valid mark")
        return get_mark(name,subject,maxmarks)
