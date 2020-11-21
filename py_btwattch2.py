import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkfd
from tkinter import messagebox
import threading, asyncio

import sys, time, datetime
from bleak import BleakClient, discover
from functools import reduce
import bisect
import csv

GATT_CHARACTERISTICS_UUID_TX = '6e400002-b5a3-f393-e0a9-e50e24dcca9e'
GATT_CHARACTERISTICS_UUID_RX = '6e400003-b5a3-f393-e0a9-e50e24dcca9e'
CMD_HEADER = bytearray.fromhex('aa')

CMD_REALTIME_MONITORING = bytearray.fromhex('08')
CMD_TURN_ON = bytearray.fromhex('a701')
CMD_TURN_OFF = bytearray.fromhex('a700')

def crc8(payload: bytearray):
    polynomial = 0x85
    def crc1(crc, times=0):
        if times >= 8:
            return crc
        else:
            if crc & 0x80:
                return crc1((crc << 1 ^ polynomial) & 0xff, times+1)
            else:
                return crc1(crc << 1, times+1)
    
    return reduce(lambda x, y: crc1(y & 0xff ^ x), payload, 0x00)

def print_measurement(voltage, current, wattage, timestamp):
    print(
        "{0},{1:.3f}W,{2:.3f}V,{3:.3f}mA".format(timestamp, wattage, voltage, current)
    )

class BTWATTCH2:
    def __init__(self, address):
        self.address = address
        self.client = None
        self.loop = asyncio.get_event_loop()

        self.services = self.loop.run_until_complete(self.setup())
        self.Tx = self.services.get_characteristic(GATT_CHARACTERISTICS_UUID_TX)
        self.Rx = self.services.get_characteristic(GATT_CHARACTERISTICS_UUID_RX)
        self.enable_notify()

        self.callback = print_measurement

    async def setup(self):
        self.client = BleakClient(self.address)
        
        await self.client.connect()
        return await self.client.get_services()
    
    def enable_notify(self):
        async def _enable_notify():
            await self.client.start_notify(self.Rx, self.format_message())
        
        return self.loop.run_until_complete(_enable_notify())

    def disable_notify(self):
        async def _disable_notify():
            await self.client.stop_notify(self.Rx)
        
        return self.loop.run_until_complete(_disable_notify())

    def cmd(self, payload: bytearray):
        header = CMD_HEADER
        pld_length = len(payload).to_bytes(2, 'big')
        return header + pld_length + payload + crc8(payload).to_bytes(1, 'big')

    def write(self, payload: bytearray):
        async def _write(payload):
            command = self.cmd(payload)
            await self.client.write_gatt_char(self.Tx, command, True)
            
        if self.loop.is_running():
            return self.loop.create_task(_write(payload))
        else:
            return self.loop.run_until_complete(_write(payload))

    def set_rtc(self):
        time.sleep(1 - datetime.datetime.now().microsecond/1e6)

        d = datetime.datetime.now().timetuple()
        payload = 0x01, d.tm_sec, d.tm_min, d.tm_hour, d.tm_mday, d.tm_mon-1, d.tm_year-1900, d.tm_wday
        self.write(bytearray(payload))

    def on(self):
        self.write(CMD_TURN_ON)
        
    def off(self):
        self.write(CMD_TURN_OFF)

    def measure(self):
        self.write(CMD_REALTIME_MONITORING)

    def format_message(self):
        buffer = bytearray()
        def _format_message(sender: int, data: bytearray):
            if data[0] == CMD_HEADER[0]:
                nonlocal buffer
                buffer = buffer + data
            else:
                data = buffer + data
                if data[3] == 0x08:
                    voltage = int.from_bytes(data[5:11], 'little') / (16**6)
                    current = int.from_bytes(data[11:17], 'little') / (32**6) * 1000
                    wattage = int.from_bytes(data[17:23], 'little') / (16**6)
                    timestamp = datetime.datetime(1900+data[28], data[27]+1, *data[26:22:-1])
                    self.callback(voltage, current, wattage, timestamp)
                else:
                    pass    # to be implemented
                
                buffer.clear()
        
        return _format_message

class main(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.master.title('RS-BTWATTCH2')

        self.wattchecker = None
        self.discover_wattcheker()

        self.tree = None
        self.columns = None
        self.selected_column = 0
        self.column_reversed = True

        self.started = threading.Event()
        self.running = True
        thread = threading.Thread(target=self._measure_thread)
        thread.start()
        self.master.protocol('WM_DELETE_WINDOW', self._kill_measure)

    def discover_wattcheker(self):
        def tt(selected):
            frame_device_list.destroy()
            ordinal = selected.get()
            bdaddr = list_wattchecker[ordinal].address
            self.setup_wattcheker(bdaddr)
        
        frame_device_list = tk.Frame(self.master)
        frame_device_list.grid(sticky=tk.NSEW)
        self.master.resizable(False, False)

        ble_devices = asyncio.get_event_loop().run_until_complete(discover())
        list_wattchecker = [d for d in ble_devices if 'BTWATTCH2' in d.name]

        if list_wattchecker:
            selected = tk.IntVar()
            for i in range(len(list_wattchecker)):
                ttk.Radiobutton(frame_device_list, value=i, variable=selected, text=list_wattchecker[i]).pack()
            
            button = ttk.Button(frame_device_list, text='connect', command=lambda: tt(selected))
            button.pack(anchor=tk.CENTER)
        else:
            messagebox.showerror('RS-BTWATTCH2', 'Device not found')
            sys.exit(0)

    def setup_wattcheker(self, bdaddr):
        self.wattchecker = BTWATTCH2(bdaddr)
        self.wattchecker.set_rtc()
        self.wattchecker.callback = self.add_row
        self._create_widgets()

    def add_row(self, voltage, current, wattage, timestamp):
        measurement = timestamp, round(wattage,3), int(current), round(voltage,2)
        curr_col = [self.tree.set(k, self.selected_column) for k in self.tree.get_children('')]
        to_insert = measurement[self.selected_column]
        
        if self.column_reversed:
            reversed_index = bisect.bisect_right(curr_col[::-1], str(to_insert))
            index = len(curr_col) - reversed_index
        else:
            index = bisect.bisect_left(curr_col, str(to_insert))

        self.tree.insert('', index=index, values=measurement)

    def sort_column(self, treeview, column):
        is_reverse = not self.column_reversed

        l = [(treeview.set(k, column), k) for k in treeview.get_children('')]
        l.sort(key=lambda x: x[0], reverse=is_reverse)

        for index, (_, item_id) in enumerate(l):
            treeview.move(item_id, '', index)

        self.selected_column = column
        self.column_reversed = is_reverse

    def _kill_measure(self):
        self.running = False
        self.started.set()
        self.master.destroy()

    def _measure_thread(self):
        self.started.wait()
        while self.running:
            if self.started.is_set():
                self.wattchecker.measure()

                ms = datetime.datetime.now().microsecond
                self.wattchecker.loop.run_until_complete(asyncio.sleep(1.05 - ms/1e6))
            else:
                self.started.wait(3)
        
    def _create_widgets(self):
        self.master.resizable(True, True)
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)
        self._create_button()
        self._place_treeview()
        self._set_columns()
      
    def _create_button(self):
        frame_button = tk.Frame(self.master)
        frame_button.grid(sticky=tk.NSEW)

        button1 = ttk.Button(frame_button, text='ON', width=5)
        button1.bind('<Button-1>', lambda event: self.wattchecker.on())
        button1.pack(anchor=tk.NW, side=tk.LEFT)

        button2 = ttk.Button(frame_button, text='OFF', width=5)
        button2.bind('<Button-1>', lambda event: self.wattchecker.off())
        button2.pack(anchor=tk.NW, side=tk.LEFT)

        button3 = ttk.Button(frame_button, text='measure', default=tk.ACTIVE)
        button3.bind('<Button-1>', lambda event: self._measure_btn_clicked(button3))
        button3.pack(anchor=tk.NW, side=tk.LEFT)

        button4 = ttk.Button(frame_button, text='clear')
        button4.bind('<Button-1>', lambda event: self.tree.delete(*self.tree.get_children()))
        button4.pack(anchor=tk.NE, side=tk.RIGHT)

        button5 = ttk.Button(frame_button, text='save as')
        button5.bind('<Button-1>', lambda event: self._save_csv())
        button5.pack(anchor=tk.NE, side=tk.RIGHT)

    def _save_csv(self):
        out = []
        for child in self.tree.get_children(''):
            row = self.tree.item(child, 'values')
            out.append(row)
        
        fname = tkfd.asksaveasfilename(filetypes=[('CSV File', '*.csv'),('', '*.*')], defaultextension='.csv', initialdir='./')
        if fname:
            with open(fname, mode='w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['datetime', 'voltage[V]', 'current[mA]', 'voltage[V]'])
                    writer.writerows(out)
        
        return 'break'
        
    def _measure_btn_clicked(self, button):
        if self.started.is_set():
            button.configure(text='measure')
            self.started.clear()
        else:
            button.configure(text='stop')
            self.started.set()

    def _place_treeview(self):
        frame_treeview = tk.Frame(self.master)
        frame_treeview.grid(sticky=tk.NSEW)
        frame_treeview.columnconfigure(0, weight=1)
        frame_treeview.rowconfigure(0, weight=1)

        columns = (0, 1, 2, 3)
        self.columns = columns
        self.tree = ttk.Treeview(frame_treeview, columns=columns, show='headings', height=25)
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)

        vscrollbar = ttk.Scrollbar(frame_treeview, orient=tk.VERTICAL, command=self.tree.yview)
        vscrollbar.grid(row=0, column=1, sticky=tk.N+tk.S)
        hscrollbar = ttk.Scrollbar(frame_treeview, orient=tk.HORIZONTAL, command=self.tree.xview)
        hscrollbar.grid(row=1, column=0, sticky=tk.E+tk.W)
        self.tree.configure(yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set)

    def _set_columns(self):
        self.tree.column(self.columns[0], width=150, minwidth=100, stretch=tk.NO)
        self.tree.column(self.columns[1], width=100, minwidth=100, stretch=tk.NO)
        self.tree.column(self.columns[2], width=100, minwidth=100, stretch=tk.NO)
        self.tree.column(self.columns[3], width=100, minwidth=100)

        self.tree.heading(self.columns[0], text="datetime", command=lambda: self.sort_column(self.tree, self.columns[0]))
        self.tree.heading(self.columns[1], text="wattage[W]", command=lambda: self.sort_column(self.tree, self.columns[1]))
        self.tree.heading(self.columns[2], text="current[mA]", command=lambda: self.sort_column(self.tree, self.columns[2]))
        self.tree.heading(self.columns[3], text="voltage[V]", command=lambda: self.sort_column(self.tree, self.columns[3]))

if __name__ == "__main__":
    base = tk.Tk()
    main(base)
    base.mainloop()