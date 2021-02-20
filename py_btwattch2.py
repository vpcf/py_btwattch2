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

GATT_CHARACTERISTIC_UUID_TX = '6e400002-b5a3-f393-e0a9-e50e24dcca9e'
GATT_CHARACTERISTIC_UUID_RX = '6e400003-b5a3-f393-e0a9-e50e24dcca9e'
GATT_CHARACTERISTIC_DEVICE_NAME = '00002A24-0000-1000-8000-00805F9B34FB'
CMD_HEADER = bytearray.fromhex('aa')

PAYLOAD_TIMER = bytearray.fromhex('01')
PAYLOAD_TURN_ON = bytearray.fromhex('a701')
PAYLOAD_TURN_OFF = bytearray.fromhex('a700')
PAYLOAD_REALTIME_MONITORING = bytearray.fromhex('08')

def crc8(payload: bytearray):
    POLYNOMIAL = 0x85
    MSBIT = 0x80
    def crc1(crc, step=0):
        if step >= 8:
            return crc & 0xff
        elif crc & MSBIT :
            return crc1(crc << 1 ^ POLYNOMIAL, step+1)
        else:
            return crc1(crc << 1, step+1)
    
    return reduce(lambda x, y: crc1(y ^ x), payload, 0x00)

def print_measurement(voltage, current, wattage, timestamp):
    print('{{"datetime":"{0}", "wattage":{1:.3f}, "voltage":{2:.3f}, "current":{3:.3f}}}'.format(timestamp, wattage, voltage, current))

class BTWATTCH2:
    def __init__(self, address):
        self.address = address
        self.client = None
        self.loop = asyncio.get_event_loop()

        self.services = self.loop.run_until_complete(self.setup())
        self.Tx = self.services.get_characteristic(GATT_CHARACTERISTIC_UUID_TX)
        self.Rx = self.services.get_characteristic(GATT_CHARACTERISTIC_UUID_RX)
        self.char_device_name = self.services.get_characteristic(GATT_CHARACTERISTIC_DEVICE_NAME)
        self.enable_notify()

        self.callback = print_measurement

    @property
    def model_number(self):
        model_number_bytearray = asyncio.run(self.client.read_gatt_char(self.char_device_name))
        return model_number_bytearray.decode()

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
        pld_length = len(payload).to_bytes(2, 'big')
        return CMD_HEADER + pld_length + payload + crc8(payload).to_bytes(1, 'big')

    def write(self, payload: bytearray):
        async def _write(payload):
            command = self.cmd(payload)
            await self.client.write_gatt_char(self.Tx, command, True)
            
        if self.loop.is_running():
            return self.loop.create_task(_write(payload))
        else:
            return self.loop.run_until_complete(_write(payload))

    def set_timer(self):
        time.sleep(1 - datetime.datetime.now().microsecond/1e6)

        d = datetime.datetime.now().timetuple()
        payload = (
            PAYLOAD_TIMER[0], 
            d.tm_sec, d.tm_min, d.tm_hour, 
            d.tm_mday, d.tm_mon-1, d.tm_year-1900, 
            d.tm_wday
        )
        self.write(bytearray(payload))

    def on(self):
        self.write(PAYLOAD_TURN_ON)
        
    def off(self):
        self.write(PAYLOAD_TURN_OFF)

    def measure(self):
        self.write(PAYLOAD_REALTIME_MONITORING)
        ms = datetime.datetime.now().microsecond
        self.loop.run_until_complete(asyncio.sleep(1.05 - ms/1e6))

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
    def __init__(self, master, wattchecker):
        super().__init__(master)
        self.master = master
        self.tree = None
        self.headings = ('datetime', 'wattage[W]', 'current[mA]', 'voltage[V]')

        self.wattchecker = wattchecker
        self.master.title(self.wattchecker.model_number)
        self.wattchecker.callback = self.add_row

        self._create_widgets()
        self.active_column = self.headings[0]
        self.is_ascending = False
        
        self.started = threading.Event()
        self.running = True
        thread = threading.Thread(target=self._measure_thread)
        thread.start()
        
        self.master.protocol('WM_DELETE_WINDOW', self._kill_app)

    def locate_insertion_position(self, measurement):
        active_col = [self.tree.set(k, self.active_column) for k in self.tree.get_children('')]
        new_col_element = measurement[self.headings.index(self.active_column)]
        
        if self.active_column == self.headings[0]:
            lst = active_col
            element = str(new_col_element)
        else:
            lst = [float(f) for f in active_col]
            element = float(new_col_element)
            
        if self.is_ascending:
            return bisect.bisect_left(lst, element)
        else:
            return len(lst) - bisect.bisect_right(lst[::-1], element)

    def add_row(self, voltage, current, wattage, timestamp):
        measurement = timestamp, round(wattage, 3), int(current), round(voltage, 2)
        position_to_insert = self.locate_insertion_position(measurement)
        self.tree.insert('', index=position_to_insert, values=measurement)

    def sort_column(self, treeview, heading):
        self.is_ascending = not self.is_ascending
        self.active_column = heading
        
        if self.active_column == self.headings[0]:
            func = lambda x: x[0]
        else:
            func = lambda x: float(x[0])
        
        l = [(treeview.set(k, heading), k) for k in treeview.get_children('')]
        l.sort(key=func, reverse=not self.is_ascending)

        for index, (_, item_id) in enumerate(l):
            treeview.move(item_id, '', index)

    def _kill_app(self):
        self.running = False
        self.started.set()
        self.master.destroy()

    def _measure_thread(self):
        self.started.wait()
        while self.running:
            if self.started.is_set():
                self.wattchecker.measure()
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

        ttk.Style().layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
        self.tree = ttk.Treeview(frame_treeview, style='Treeview', columns=self.headings, show='headings', height=25)
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)

        vscrollbar = ttk.Scrollbar(frame_treeview, orient=tk.VERTICAL, command=self.tree.yview)
        vscrollbar.grid(row=0, column=1, sticky=tk.N+tk.S)
        hscrollbar = ttk.Scrollbar(frame_treeview, orient=tk.HORIZONTAL, command=self.tree.xview)
        hscrollbar.grid(row=1, column=0, sticky=tk.E+tk.W)
        self.tree.configure(yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set)

    def _set_columns(self):
        self.tree.column(self.headings[0], width=150, minwidth=100, stretch=tk.NO)
        self.tree.column(self.headings[1], width=100, minwidth=100, stretch=tk.NO)
        self.tree.column(self.headings[2], width=100, minwidth=100, stretch=tk.NO)
        self.tree.column(self.headings[3], width=100, minwidth=100)

        self.tree.heading(self.headings[0], text=self.headings[0], command=lambda: self.sort_column(self.tree, self.headings[0]))
        self.tree.heading(self.headings[1], text=self.headings[1], command=lambda: self.sort_column(self.tree, self.headings[1]))
        self.tree.heading(self.headings[2], text=self.headings[2], command=lambda: self.sort_column(self.tree, self.headings[2]))
        self.tree.heading(self.headings[3], text=self.headings[3], command=lambda: self.sort_column(self.tree, self.headings[3]))

def setup_btwattch2(bdaddr):
    wattchecker = BTWATTCH2(bdaddr)
    wattchecker.set_timer()

    base = tk.Tk()
    main(base, wattchecker)

    base.mainloop()

def discover_btwattch2():
    ble_devices = asyncio.get_event_loop().run_until_complete(discover())
    return [d for d in ble_devices if 'BTWATTCH2' in d.name]

def device_selection_window():
    def confirm_selection(selected):
        dialog.destroy()
        ordinal = selected.get()
        bdaddr = list_wattchecker[ordinal].address
        setup_btwattch2(bdaddr)

    dialog = tk.Tk()
    dialog.resizable(False, False)
    frame_device_list = tk.Frame(dialog)
    frame_device_list.grid(sticky=tk.NSEW)

    list_wattchecker = discover_btwattch2()
    if list_wattchecker:
        selected = tk.IntVar()
        for i in range(len(list_wattchecker)):
            ttk.Radiobutton(frame_device_list, value=i, variable=selected, text=list_wattchecker[i]).pack()
        
        button = ttk.Button(frame_device_list, text='connect', command=lambda: confirm_selection(selected))
        button.pack(anchor=tk.CENTER)
    else:
        messagebox.showerror('RS-BTWATTCH2', 'Device not found')
        sys.exit(0)
    
    dialog.mainloop()

if __name__ == '__main__':
    device_selection_window()