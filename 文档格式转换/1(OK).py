import sys, os
import comtypes.client


def init_word():
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 1
    return word


def convert_files_in_folder(word, folder):
    files = os.listdir(folder)
    # 将所有以.doc(x)结尾的文件加入pwd path
    wordfiles = [f for f in files if f.endswith((".doc", ".docx"))]
    for wordfile in wordfiles:
        print(wordfile)
        if wordfile + '.pdf' in files:
            break
        fullpath = os.path.join(pwd, wordfile)
        word_to_pdf(word, fullpath, fullpath)


def word_to_pdf(word, input_file_name, output_file_name, formatType=17):
    if output_file_name[-3:] != 'pdf':
        output_file_name = output_file_name + ".pdf"
    print(input_file_name)
    deck = word.Documents.Open(input_file_name)
    deck.SaveAs(output_file_name, formatType)
    deck.Close()


if __name__ == "__main__":
    word = init_word()
    pwd = os.getcwd()
    print(pwd)
    convert_files_in_folder(word, pwd)
    word.Quit()