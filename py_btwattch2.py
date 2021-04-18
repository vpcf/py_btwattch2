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

UART_TX_UUID = '6e400002-b5a3-f393-e0a9-e50e24dcca9e'
UART_RX_UUID = '6e400003-b5a3-f393-e0a9-e50e24dcca9e'
DEVICE_NAME_UUID = '00002A24-0000-1000-8000-00805F9B34FB'
CMD_HEADER = bytearray.fromhex('aa')

ID_TIMER = bytearray.fromhex('01')
ID_TURN_ON = bytearray.fromhex('a701')
ID_TURN_OFF = bytearray.fromhex('a700')
ID_ENERGY_USAGE = bytearray.fromhex('08')

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
    
    return reduce(lambda crc, input_byte: crc1(input_byte ^ crc), payload, 0x00)

def print_measurement(timestamp, wattage, voltage, current):
    print('{{"datetime":"{0}", "wattage":{1:.3f}, "voltage":{2:.3f}, "current":{3:.3f}}}'
            .format(timestamp, wattage, voltage, current)
        )

class BTWATTCH2:
    def __init__(self, address):
        self.client = BleakClient(address)
        self.loop = asyncio.get_event_loop()

        self.services = self.loop.run_until_complete(self.setup())
        self.Tx = self.services.get_characteristic(UART_TX_UUID)
        self.Rx = self.services.get_characteristic(UART_RX_UUID)
        self.char_device_name = self.services.get_characteristic(DEVICE_NAME_UUID)
        self.loop.create_task(self.enable_notify())
        self.callback = print_measurement

    @property
    def address(self):
        return self.client.address

    @property
    def model_number(self):
        read_device_name = self.client.read_gatt_char(self.char_device_name)
        return self.loop.run_until_complete(read_device_name).decode()

    async def setup(self):
        await self.client.connect()
        return await self.client.get_services()
    
    async def enable_notify(self):
        await self.client.start_notify(self.Rx, self._cache_message())
        
    async def disable_notify(self):
        await self.client.stop_notify(self.Rx)

    def pack_command(self, payload: bytearray):
        pld_length = len(payload).to_bytes(2, 'big')
        return CMD_HEADER + pld_length + payload + crc8(payload).to_bytes(1, 'big')

    def _write(self, payload: bytearray):
        async def _write_(payload):
            command = self.pack_command(payload)
            await self.client.write_gatt_char(self.Tx, command, True)
            
        if self.loop.is_running():
            return self.loop.create_task(_write_(payload))
        else:
            return self.loop.run_until_complete(_write_(payload))

    def set_timer(self):
        time.sleep(1 - datetime.datetime.now().microsecond/1e6)

        d = datetime.datetime.now().timetuple()
        payload = (
            ID_TIMER[0], 
            d.tm_sec, d.tm_min, d.tm_hour, 
            d.tm_mday, d.tm_mon-1, d.tm_year-1900, 
            d.tm_wday
        )
        self._write(bytearray(payload))

    def on(self):
        self._write(ID_TURN_ON)

    def off(self):
        self._write(ID_TURN_OFF)

    def measure(self):
        self._write(ID_ENERGY_USAGE)
        interval = 1.05 - datetime.datetime.now().microsecond/1e6
        self.loop.run_until_complete(asyncio.sleep(interval))

    def _cache_message(self):
        buffer = bytearray()
        def _cache_message_(sender: int, value: bytearray):
            nonlocal buffer
            buffer = buffer + value
        
            if buffer[0] == CMD_HEADER[0]:
                payload_length = int.from_bytes(buffer[1:3], 'big')
                if len(buffer[3:-1]) < payload_length:
                    return
                elif len(buffer[3:-1]) == payload_length:
                    if crc8(buffer[3:]) == 0:
                        self._classify_response(buffer)

            buffer.clear()
        
        return _cache_message_

    def _classify_response(self, data):
        if data[3] == ID_ENERGY_USAGE[0]:
            measurement = self.decode_measurement(data)
            self.callback(**measurement)
        else:
            pass    # to be implemented

    def decode_measurement(self, data: bytearray):
        return {
            "voltage": int.from_bytes(data[5:11], 'little') / (16**6),
            "current": int.from_bytes(data[11:17], 'little') / (32**6) * 1000,
            "wattage": int.from_bytes(data[17:23], 'little') / (16**6),
            "timestamp": datetime.datetime(1900+data[28], data[27]+1, *data[26:22:-1]),
        }

class main(ttk.Frame):
    def __init__(self, master, wattchecker):
        super().__init__(master)
        self.master = master

        self._create_button()
        self.treeview_widget = treeview_widget(self.master)
        self._organize_widgets()

        self.wattchecker = wattchecker
        self.master.title(self.wattchecker.model_number)
        self.wattchecker.callback = self.treeview_widget.add_row
        
        self.started = threading.Event()
        self.running = True
        thread = threading.Thread(target=self._measure_thread)
        thread.start()
        
        self.master.protocol('WM_DELETE_WINDOW', self._kill_app)

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
        
    def _organize_widgets(self):
        self.master.resizable(True, True)
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)
        self.grid(row=0)
        self.treeview_widget.grid(row=1)

    def _create_button(self):
        self.grid(sticky=tk.NSEW)
        
        button1 = ttk.Button(self, text='ON', width=5)
        button1.bind('<Button-1>', lambda event: self.wattchecker.on())
        button1.pack(anchor=tk.NW, side=tk.LEFT)

        button2 = ttk.Button(self, text='OFF', width=5)
        button2.bind('<Button-1>', lambda event: self.wattchecker.off())
        button2.pack(anchor=tk.NW, side=tk.LEFT)

        button3 = ttk.Button(self, text='measure', default=tk.ACTIVE)
        button3.bind('<Button-1>', lambda event: self._measure_btn_clicked(button3))
        button3.pack(anchor=tk.NW, side=tk.LEFT)

        button4 = ttk.Button(self, text='clear')
        button4.bind('<Button-1>', lambda event: self._clear_tree())
        button4.pack(anchor=tk.NE, side=tk.RIGHT)

        button5 = ttk.Button(self, text='save as')
        button5.bind('<Button-1>', lambda event: self._save_csv())
        button5.pack(anchor=tk.NE, side=tk.RIGHT)

    def _clear_tree(self):
        self.treeview_widget.clear_tree()

    def _save_csv(self):
        out = []
        for child in self.treeview_widget.tree.get_children(''):
            row = self.treeview_widget.tree.item(child, 'values')
            out.append(row)
        
        fname = tkfd.asksaveasfilename(
            filetypes=[('CSV File', '*.csv'), ('', '*.*')], 
            defaultextension='.csv', 
            initialdir='./'
        )
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

class treeview_widget(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.headings = ('datetime', 'wattage[W]', 'current[mA]', 'voltage[V]')
        self.is_ascending = False
        self.active_column = self.headings[0]
        self.tree = self._draw_treeview()
        self._set_columns(self.tree, self.headings)

    def add_row(self, timestamp, wattage, voltage, current):
        measurement = timestamp, round(wattage, 3), int(current), round(voltage, 2)
        position_to_insert = self._locate_insertion_position(measurement)
        self.tree.insert('', index=position_to_insert, values=measurement)

    def _sort_column(self, treeview, heading):
        self.is_ascending = not self.is_ascending
        self.active_column = heading

        func = lambda x: self._convert_type_by_column(x[0])
        
        l = [(treeview.set(k, heading), k) for k in treeview.get_children('')]
        l.sort(key=func, reverse=not self.is_ascending)

        for index, (_, item_id) in enumerate(l):
            treeview.move(item_id, '', index)

    def _convert_type_by_column(self, value):
        if self.active_column == self.headings[0]:
            return str(value)
        else:
            return float(value)

    def _locate_insertion_position(self, measurement):
        active_col = [self.tree.set(k, self.active_column) for k in self.tree.get_children('')]
        new_col_element = measurement[self.headings.index(self.active_column)]
        
        lst = [self._convert_type_by_column(f) for f in active_col]
        element = self._convert_type_by_column(new_col_element)
            
        if self.is_ascending:
            return bisect.bisect_left(lst, element)
        else:
            return len(lst) - bisect.bisect_right(lst[::-1], element)

    def clear_tree(self):
        self.tree.delete(*self.tree.get_children())

    def _draw_treeview(self):
        self.grid(sticky=tk.NSEW)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        ttk.Style().layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
        tree = ttk.Treeview(self, style='Treeview', columns=self.headings, show='headings', height=25)
        tree.grid(row=0, column=0, sticky=tk.NSEW)

        vscrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=tree.yview)
        vscrollbar.grid(row=0, column=1, sticky=tk.N+tk.S)
        hscrollbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=tree.xview)
        hscrollbar.grid(row=1, column=0, sticky=tk.E+tk.W)
        tree.configure(yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set)

        return tree

    def _set_columns(self, tree, headings):
        tree.column(headings[0], width=150, minwidth=100, stretch=tk.NO)
        tree.column(headings[1], width=100, minwidth=100, stretch=tk.NO)
        tree.column(headings[2], width=100, minwidth=100, stretch=tk.NO)
        tree.column(headings[3], width=100, minwidth=100)

        tree.heading(headings[0], text=headings[0], command=lambda: self._sort_column(tree, headings[0]))
        tree.heading(headings[1], text=headings[1], command=lambda: self._sort_column(tree, headings[1]))
        tree.heading(headings[2], text=headings[2], command=lambda: self._sort_column(tree, headings[2]))
        tree.heading(headings[3], text=headings[3], command=lambda: self._sort_column(tree, headings[3]))

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