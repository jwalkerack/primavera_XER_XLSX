# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
def main():
    import streamlit as st
    st.title("File Transformation Tool")
    st.sidebar.header("Choose a Transformation")

    option = st.sidebar.selectbox("Select an option", ("XER to Excel", "Excel to XER"))

    if option == "XER to Excel":
        print ("Yes")
    elif option == "Excel to XER":
        print ("Non")


if __name__ == "__main__":
    main()




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
