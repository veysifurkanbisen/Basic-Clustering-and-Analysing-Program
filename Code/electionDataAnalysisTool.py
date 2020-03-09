from Tkinter import *
import ttk
import tkFileDialog
import xlrd
from clusters import *
from PIL import ImageTk, Image

class Election_Data(Frame):      #classic gui designing process
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.grid()
        self.parent = parent
        self.initUI()

    def initUI(self):
        self.which_button = ""

        self.app_label = Label(self, text=" " * 37 + "Election Data Analysis Tool v.1.0" + " " * 38, bg="red", fg="white", font=("", "17", "bold"))
        self.app_label.grid(row=0, column=0, columnspan=19, sticky=EW)

        self.election_button = Button(self, text="Load Election Data", height=2, width=25, command=self.add_election_data)
        self.election_button.grid(row=1, column=9, columnspan=2, pady=10)

        # Cluster Buttons
        self.district_button = Button(self, text="        Cluster Districts        ", height=2, command = self.first_button)
        self.district_button.grid(row=2, column=9, sticky=EW, padx=5)

        self.parties_button = Button(self, text="  Cluster Political Parties  ", height=2, command = self.second_button)
        self.parties_button.grid(row=2, column=10, sticky=EW, padx=5)

        # Canvas
        self.tree_canvas = Canvas(self, width=740, height=300, bg="gray", scrollregion=(0, 0, 1500, 1000))
        self.tree_canvas.configure(scrollregion=self.tree_canvas.bbox("all"))

        # Canvas Scroll Y
        self.canvas_scroll_y = Scrollbar(self, orient=VERTICAL)
        self.canvas_scroll_y.configure(command=self.tree_canvas.yview)
        self.tree_canvas.configure(yscrollcommand=self.canvas_scroll_y.set)

        # Canvas Scroll X
        self.canvas_scroll_x = Scrollbar(self, orient=HORIZONTAL)
        self.canvas_scroll_x.configure(command=self.tree_canvas.xview)
        self.tree_canvas.configure(xscrollcommand=self.canvas_scroll_x.set)

        # Cutting point Widgets

        self.percentage_label = Label(self, text="Treshold")
        self.percentage = ttk.Combobox(self, values=["%0", "%1", "%10", "%20", "%30", "%40", "%50"])


        self.run_analysis_button = Button(self, text="Refine Analysis", height=2, width=19, command = self.analyze_button)

        # clusters listbox
        self.clusterlistbox_label = Label(self, text="Clusters")
        self.cluster_listbox = Listbox(self, width=47, selectmode = MULTIPLE, height=7)
        self.cluster_listbox.config(exportselection=FALSE)

        self.listbox_scroll = Scrollbar(self, orient=VERTICAL)
        self.listbox_scroll.config(command=self.cluster_listbox.yview)
        self.cluster_listbox.config(yscrollcommand=self.listbox_scroll.set)

        self.listbox_selections = self.cluster_listbox.curselection()

    def add_election_data(self):
        self.file_path = tkFileDialog.askopenfilename()
        self.re_run_data_process(self.file_path)

    def re_run_data_process(self, path):
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)

        total_row = sheet.nrows

        row_number = 11

        self.dict = {}

        while row_number < total_row:
            ilce_value = sheet.cell_value(row_number, 2)
            total_votes = sheet.cell_value(row_number, 7)

            self.dict.setdefault(ilce_value, {})

            for i in range(9, 21, 1):
                vote_percent = 100*(sheet.cell_value(row_number, i))/total_votes
                self.dict[ilce_value].setdefault(sheet.cell_value(10, i), vote_percent)

            row_number += 1

        listboxitems = []
        for key, value in self.dict.items():
            listboxitems.append(key)
        if self.cluster_listbox.size() == 0:
            if len(self.cluster_listbox.curselection()) == 0:
                num = 0
                list = []
                for i in listboxitems:
                    list.append(i)
                    list.sort()

                while num < len(list):
                    self.cluster_listbox.insert(END, list[num])
                    num += 1

    def data_creation(self):
        self.data = {}
        self.data2 = {}
        self.list1 = []
        self.list2 = []
        for key, values in self.dict.items():
            self.list1.append(key)

            for i in values.keys():

                if i not in self.list2:
                    self.list2.append(i)

                self.data.setdefault(i, []).append(values[i])
                self.data2.setdefault(key, []).append(values[i])

        if self.which_button == "first":
            self.data_refine_1()
        if self.which_button == "second":
            self.data_refine_2()

    def data_refine_1(self):
        data = open("data.txt", "w+")
        data.write("parties")
        i = 0
        while i < len(self.list1):
            data.write('\t' + self.list1[i].encode("utf-8"))
            i+=1

        data.write("\n")

        for key in self.data.keys():
            data.write(key.encode("utf-8"))

            for percent in self.data[key]:
                data.write("\t")
                data.write(str(percent))
            data.write("\n")

        data.close()
        self.cluster_1()

    def data_refine_2(self):
        data2 = open("data2.txt", "w+")
        data2.write("districts")
        i = 0
        while i < len(self.list2):
            data2.write('\t' + self.list2[i].encode("utf-8"))
            i += 1

        data2.write("\n")

        for key in self.data2.keys():
            data2.write(key.encode("utf-8"))

            for percent in self.data2[key]:
                data2.write("\t")
                data2.write(str(percent))
            data2.write("\n")

        data2.close()
        self.cluster_2()

    def cluster_1(self):
        ids, codes, datas = readfile("data.txt")
        clust = hcluster(datas)
        kclust = kcluster(datas)

        drawdendrogram(clust, ids, jpeg="cls.jpg")

        self.image1 = Image.open("cls.jpg")
        self.photo = ImageTk.PhotoImage(self.image1)

    def cluster_2(self):
        ids, codes, datas = readfile("data2.txt")
        clust = hcluster(datas)
        kclust = kcluster(datas)

        drawdendrogram(clust, ids, jpeg="cls2.jpg")

        self.image2 = Image.open("cls2.jpg")
        self.photo2 = ImageTk.PhotoImage(self.image2)

    def first_button(self):
        self.tree_canvas.delete("all")
        self.which_button = ""
        self.which_button += "first"
        self.tree_canvas.grid(row=3, column=0, columnspan=19)

        self.canvas_scroll_y.grid(column=17, row=3, sticky=NS + W)
        self.canvas_scroll_x.grid(column=3, row=4, sticky=EW + N, columnspan=14)

        self.clusterlistbox_label.grid(row=5, column=7, columnspan=3,  sticky=EW)
        self.cluster_listbox.grid(row=6, column=7, columnspan=3)
        self.listbox_scroll.grid(row=6, column=7, columnspan=3, sticky=NS+E)
        self.percentage_label.grid(row=6, column=10, sticky=N)
        self.percentage.grid(row=6, column=10)
        self.percentage.current(0)
        self.run_analysis_button.grid(row=6, column=11, sticky=E)

        self.data_creation()

        self.tree_canvas.create_image(0, 0, image=self.photo, anchor=NW)
        self.tree_canvas.grid(row=3, column=0, columnspan=19)


    def second_button(self):
        self.tree_canvas.delete("all")
        self.which_button = ""
        self.which_button += "second"
        self.tree_canvas.grid(row=3, column=0, columnspan=19)

        self.canvas_scroll_y.grid(column=17, row=3, sticky=NS + W)
        self.canvas_scroll_x.grid(column=3, row=4, sticky=EW + N, columnspan=14)

        self.clusterlistbox_label.grid(row=5, column=7, columnspan=3, sticky=EW)
        self.cluster_listbox.grid(row=6, column=7, columnspan=3)
        self.listbox_scroll.grid(row=6, column=7, columnspan=3, sticky=NS + E)
        self.percentage_label.grid(row=6, column=10, sticky=N)
        self.percentage.grid(row=6, column=10)
        self.percentage.current(0)
        self.run_analysis_button.grid(row=6, column=11, sticky=E)

        self.data_creation()

        self.tree_canvas.create_image(0, 0, image=self.photo2, anchor=NW)
        self.tree_canvas.grid(row=3, column=0, columnspan=19)

    def analyze_button(self):

        keep = []
        to_be_included = self.cluster_listbox.curselection()
        for i in to_be_included:
            keep.append(self.cluster_listbox.get(i))

        treshold = self.percentage.get()[1:]

        self.cluster_listbox.selection_clear(0, END)

        if len(keep) > 1:
            for keys, values in self.dict.items():
                if keys not in keep:
                    del self.dict[keys]

        for keys, values in self.dict.items():
            for i in values.keys():
                if int(values[i]) < int(treshold[0:]):
                    values[i] = 0

        if self.which_button == "first":

            self.data_creation()
            self.tree_canvas.delete("all")
            self.tree_canvas.create_image(0, 0, image=self.photo, anchor=NW)
            self.tree_canvas.grid(row=3, column=0, columnspan=19)
            self.re_run_data_process(self.file_path)

        if self.which_button == "second":

            self.data_creation()
            self.tree_canvas.delete("all")
            self.tree_canvas.create_image(0, 0, image=self.photo2, anchor=NW)
            self.tree_canvas.grid(row=3, column=0, columnspan=19)
            self.re_run_data_process(self.file_path)

root = Tk()
root.title("Election Data Analysis Engine")
root.geometry("817x800+0+0")
app = Election_Data(root)
app.pack(fill=BOTH, expand=True)
root.mainloop()