import sys
import subprocess
import tkinter as tk
from tkinter import ttk
import time
from datetime import datetime
import threading
import queue

VERSION = "0.9"
MAX_AGE = 30
SKED_START = "=========== SKCC Sked Page ============"


class SkimmerWrapper(threading.Thread):
    def __init__(self, queue_param):
        super().__init__()
        print("in init for skimmer wrapper")
        self.alive = True
        self.queue = queue_param

        if sys.platform == "win32":
            self.stream = subprocess.Popen(["skcc_skimmer.exe"],
                                           stdout=subprocess.PIPE,
                                           encoding='UTF-8',
                                           bufsize=1)
        else:
            self.stream = subprocess.Popen([sys.executable, "-u", "skcc_skimmer.py"],
                                           stdout=subprocess.PIPE,
                                           encoding='UTF-8',
                                           bufsize=1,
                                           text=True)

    def run(self):
        while self.alive:
            self.process_incoming()

    def stop(self):
        self.stream.terminate()
        self.alive = False

    def process_incoming(self):
        line = self.stream.stdout.readline().strip()
        if line:
            self.process_line(line)

    @staticmethod
    def line_to_spot(line: str, is_rbn_spot: bool):
        zulu = line[0:5]
        # flag = line[5]  # comes in as '+' for new RBN spots, but we're already filtering before this
        call = line[6:12].strip()
        skcc_nr = line[14:19].strip()
        skcc_level = line[20]
        skcc = skcc_nr + " " + skcc_level
        name = line[25:35].strip()
        qth = line[35:38].strip()
        freq = ''
        you_need = ''
        status_comment = ''
        wpm = None
        if is_rbn_spot:
            # RBN spots can come in like either:
            # 1612Z+K4AHO  ( 1235 T    Jim        FL) on  14059.9 by W3RGA(660mi, 11dB); YOU need them for Tx4
            # 1613Z+K7QB   ( 5733 S    Bob        IN); Last spotted 2 minutes ago on 7058.0; YOU need them for Tx4
            # combination of both RBN+sked - process as RBN spot and put status in need field:
            # 1756Z+W2TJ   ( 9330 Tx6  Tom        NY); Last spotted 42 seconds ago on 7116.0; YOU need them for Tx4; STATUS: `7.116
            if "Last spotted" in line:
                ago_idx = line.find("ago on ")
                freq = line[ago_idx + 7: line.find(';', ago_idx)].strip()
            else:
                freq = line[42:51].strip()
                # for a modified version of skcc_skimmer.py that passes along WPM too
                wpm_idx = line.find(" WPM)")
                if wpm_idx != -1:
                    wpm_start_idx = line.rfind("(", 0, wpm_idx)
                    wpm = line[wpm_start_idx + 1: wpm_idx]

        # sked page spots can come in like
        # 1642Z KA3LOC (  660 Sx6  Ric        KS); YOU need them for Tx4
        # 1642Z+KA2FIR ( 3377 T    Mike       NJ); YOU need them for Tx4; STATUS: Need AK on 80M for #50 on LOTW, etc. prop.kc2g.com
        # 1612Z AB4PP  (   32 Sx2  John-Paul  NC); YOU need them for BRAG,C,T,WAS,WAS-C,WAS-T,WAS-S,P(new +32); THEY need you for Sx3; STATUS: looking for /AF& OC

        for part in line.split(';'):
            part = part.strip()
            if part.startswith('YOU need them for '):
                you_need = part[len('YOU need them for '):]
            if part.startswith('THEY need you for '):
                # not doing anything with this
                they_need = part[len('THEY need you for '):]
            if part.startswith('STATUS: '):
                status_comment = part[len('STATUS: '):]

        return Spot(is_rbn_spot, zulu, call, skcc, name, qth, freq, you_need, status_comment, wpm)

    def process_line(self, line):
        # line = line.decode('ascii').strip()
        # can't figure out how to get rid of these bell chars earlier
        if ord(line[0]) == 7:
            line = line[1:]
        if ord(line[-1]) == 7:
            line = line[0:-1]
        print("|" + line + "|")

        # show progress lines if they come in
        if line[0] == '.':
            self.queue.put(line)
            return

        if line == SKED_START:
            self.queue.put(line)
            return

        if len(line) > 4 and line[4] != 'Z':
            pass  # not an rbn or sked line (they start like "1858Z"), skip it
        elif (") on " in line and " by " in line) or ("Last spotted " in line and " on " in line):
            if "+" in line:
                # print("RBN:", line)
                spot = self.line_to_spot(line, True)
                print("Parsed RBN spot", spot)
                self.queue.put(spot)
            else:
                # print("RBo:", line)
                pass
        elif " need " in line:
            # print("SKD:", line)
            spot = self.line_to_spot(line, False)
            print("Parsed sked spot", spot)
            self.queue.put(spot)
        else:
            print("?? Unexpected line", line)
        self.queue.put(line)


class Spot:
    def __init__(self, is_rbn_spot, zulu, call, skcc, name, qth, freq, need, comment, wpm=None):
        self.is_rbn_spot = is_rbn_spot
        self.zulu = zulu
        self.call = call
        self.skcc = skcc
        self.name = name
        self.qth = qth
        self.freq = freq
        self.need = need
        self.comment = comment
        self.wpm = wpm

    def __str__(self):
        rbn_text = "RBN" if self.is_rbn_spot else "SKED"
        return " ".join([rbn_text, self.zulu, self.call, self.skcc,
                         self.name, self.qth, self.freq, self.need, self.comment])


class GridView:
    def __init__(self):
        self.sked_spots = []
        self.rbn_spots = []

        self.root = tk.Tk()
        self.root.title("SKCC Skimmer")
        tk.Label(self.root, text='RBN spots').pack()
        column_labels = ('Age', 'Call', 'SKCC', 'Name', 'QTH', 'Frequency', 'Note')
        column_widths = ('40', '70', '70', '100', '40', '400', '170')
        ac = ('a', 'b', 'c', 'd', 'e', 'f', 'g')

        frame_top = tk.Frame(self.root)
        self.tv1 = ttk.Treeview(frame_top, columns=ac, show='headings', height=7)
        self.setup_headers(ac, self.tv1, column_labels, column_widths)

        verscrlbar1 = ttk.Scrollbar(frame_top, orient="vertical", command=self.tv1.yview)
        verscrlbar1.pack(side='right', fill='both')

        self.tv1.pack(fill=tk.Y, expand=1)
        self.tv1.configure(yscrollcommand=verscrlbar1.set)

        frame_top.pack(fill=tk.Y, expand=1)

        tk.Label(self.root).pack()  # extra space
        tk.Label(self.root, text='Sked Page').pack()

        column_labels = ('Age', 'Call', 'SKCC', 'Name', 'QTH', 'Status', 'Need')

        frame_bottom = tk.Frame(self.root)
        self.tv2 = ttk.Treeview(frame_bottom, columns=ac, show='headings', height=7)
        self.setup_headers(ac, self.tv2, column_labels, column_widths)

        verscrlbar2 = ttk.Scrollbar(frame_bottom, orient="vertical", command=self.tv2.yview)
        verscrlbar2.pack(side='right', fill='both')

        self.tv2.pack(fill=tk.Y, expand=1)
        self.tv1.configure(yscrollcommand=verscrlbar1.set)

        frame_bottom.pack(fill=tk.Y, expand=1)

        self.last_updated_var = tk.StringVar(value="Last updated:")
        tk.Label(self.root, textvar=self.last_updated_var).pack()

        self.feedback_var = tk.StringVar(value="Starting...")
        tk.Label(self.root, textvar=self.feedback_var).pack()

        self.root.protocol("WM_DELETE_WINDOW", self.cleanup)

        self.queue = queue.Queue()
        self.skimmer_wrapper = SkimmerWrapper(self.queue)
        self.skimmer_wrapper.start()
        self.next_processing = self.root.after(100, self.process_queue)

    def start(self):
        print("starting main GridView loop")
        self.root.mainloop()

    def process_queue(self):
        try:
            popped = self.queue.get_nowait()
            if isinstance(popped, Spot):
                self.add_spot(popped)
            else:
                if popped == SKED_START:
                    self.clear_sked_spots_list()
                else:
                    self.feedback(popped)
        except queue.Empty:
            pass
        self.next_processing = self.root.after(100, self.process_queue)

    def show_last_updated(self):
        self.last_updated_var.set("Updated at " + time.strftime("%H:%M:%S"))

    def feedback(self, text):
        # print(text)
        self.feedback_var.set(text[:100])

    def remove_old(self, spots):
        spots[:] = [spot for spot in spots if self.spot_age(spot) <= MAX_AGE]

    def clear_sked_spots_list(self):
        self.sked_spots.clear()
        print("Cleared sked spots list variable")

    def add_spot(self, spot):
        print("adding spot to table:", spot)
        if spot.is_rbn_spot:
            spot_list = self.rbn_spots
        else:
            spot_list = self.sked_spots

        for cur_spot in spot_list:
            if cur_spot.call == spot.call:
                spot_list.remove(cur_spot)
        spot_list.append(spot)

        self.remove_old(self.rbn_spots)
        self.remove_old(self.sked_spots)

        self.fill_grid()

    # TODO update age on a schedule, not only when there's a new line

    def fill_grid(self):
        self.fill_treeview(self.tv1, self.rbn_spots)
        self.fill_treeview(self.tv2, self.sked_spots)

        self.feedback("Finished lookup")
        self.show_last_updated()

    def spot_age(self, spot):
        return self.spot_age_mins(spot.zulu)

    def fill_treeview(self, tv, spots):
        tv.delete(*tv.get_children())
        spots.sort(key=self.spot_age)
        for index, spot in enumerate(spots):
            tv.insert('', 'end', values=self.get_row_for_table(spot), iid=index)

    def get_row_for_table(self, spot):
        if spot.freq:
            freq_wpm = spot.freq
            if spot.wpm:
                freq_wpm = f"{spot.freq:>14} ({spot.wpm} WPM)"  # adding spaces to left for alignment in UI
            return [self.spot_age_mins(spot.zulu), spot.call, spot.skcc, spot.name, spot.qth,
                    freq_wpm, spot.need]
        else:
            return [self.spot_age_mins(spot.zulu), spot.call, spot.skcc, spot.name, spot.qth,
                    spot.comment, spot.need]

    @staticmethod
    def spot_age_mins(zulu):
        zulu_hours = int(zulu[0:2])
        zulu_mins = int(zulu[2:4])
        now_zulu = datetime.utcnow()
        now_hours = now_zulu.hour
        now_mins = now_zulu.minute
        if now_hours < zulu_hours:
            now_hours += 24  # should be enough to handle pre/post 0000z
        return (now_hours - zulu_hours) * 60 + (now_mins - zulu_mins)

    @staticmethod
    def setup_headers(ac, tv, column_labels, column_widths):
        for i in range(len(column_labels)):
            tv.column(ac[i], width=column_widths[i], anchor=tk.CENTER)
            tv.heading(ac[i], text=column_labels[i])
        # tv.pack()

    def cleanup(self):
        print("cleanup")
        self.root.after_cancel(self.next_processing)
        self.root.destroy()
        self.skimmer_wrapper.stop()
        exit(0)


if __name__ == '__main__':
    print("Starting skimmer gui wrapper version", VERSION)
    grid_view = GridView()
    grid_view.start()
