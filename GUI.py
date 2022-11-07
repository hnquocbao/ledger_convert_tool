import queue
import threading
import time
import PySimpleGUI as sg
from app import converter


# ############################# User callable CPU intensive code #############################
# Put your long running code inside this "wrapper"
# NEVER make calls to PySimpleGUI from this thread (or any thread)!
# Create one of these functions for EVERY long-running call you want to make

def long_function_wrapper(work_id, gui_queue, filename):
    # LOCATION 1
    # this is our "long running function call"
    # time.sleep(10)  # sleep for a while as a simulation of a long-running computation

    try:
        if filename:
            converter(filename)
            gui_queue.put('{} ::: done'.format(work_id))
    except Exception as e:
        gui_queue.put('{0} - error: {1}'.format(work_id, e))
        # x = 0
        # while True:
        #     print(x)
        #     time.sleep(0.5)
        #     x = x + 1
        #     if x == 5:
        #         break
        # at the end of the work, before exiting, send a message back to the GUI indicating end
        
        # at this point, the thread exits
    return


def the_gui():


    gui_queue = queue.Queue()  # queue used to communicate between the gui and long-running code

    layout = [[sg.Text('')],
            #   [sg.Text('This is a Test.', size=(25, 1), key='_OUTPUT_')],
            #   [sg.Text(size=(25, 1), key='_OUTPUT2_')],
                [sg.Text("Choose a file: "),
                sg.Input(key="-IN-"),
                sg.FileBrowse(file_types=(("ALL Excel Files", "*.xlsx"), ("ALL Files", "*.*"), ))],
              [sg.Button('Go'), sg.Button('Exit')], ]

    window = sg.Window('Convert Tool', layout, size=(600,150))

    # --------------------- EVENT LOOP ---------------------
    work_id = 0
    filename = ""
    while True:
        event, values = window.Read(timeout=100)  # wait for up to 100 ms for a GUI event

        if event is None or event == 'Exit':
            # sg.PopupAnimated(None)
            break
        if event == 'Go':  # clicking "Go" starts a long running work item by starting thread
            # window.Element('_OUTPUT_').Update('Starting long work %s' % work_id)
            # LOCATION 2
            # STARTING long run by starting a thread
            filename = values['-IN-']
            thread_id = threading.Thread(target=long_function_wrapper, args=(work_id, gui_queue,filename), daemon=True)
            thread_id.start()
            # for i in range(200000):

            work_id = work_id + 1 if work_id < 19 else 0

            # while True:
            # if message == None:
            # break
        # --------------- Read next message coming in from threads ---------------
        try:
            message = gui_queue.get_nowait()  # see if something has been posted to Queue
        except queue.Empty:  # get_nowait() will get exception when Queue is empty
            message = None  # nothing in queue so do nothing
        # if message received from queue, then some work was completed
        if message is not None:
            # LOCATION 3
            # this is the place you would execute code at ENDING of long running task
            # You can check the completed_work_id variable to see exactly which long-running function completed
            # completed_work_id = int(message[:message.index(' :::')])
            # window.Element('_OUTPUT2_').Update('Finished long work %s' % completed_work_id)
            work_id -= 1
            if not work_id:
                if 'error' in message:
                    ex = message.split(':')
                    sg.PopupAnimated(None)
                    sg.popup_error_with_traceback(f'An error happened.  Here is the info:', ex[1])
                else:
                    sg.PopupAnimated(None)
                    sg.Popup('Done.')

            

        if work_id:
            sg.PopupAnimated(sg.DEFAULT_BASE64_LOADING_GIF, background_color='white', time_between_frames=100)

        # window['_GIF_'].update_animation(sg.DEFAULT_BASE64_LOADING_GIF, time_between_frames=100)
        # window.read(timeout = 1000)

    # if user exits the window, then close the window and exit the GUI func
    window.Close()

############################# Main #############################

if __name__ == '__main__':
    the_gui()
    print('Exiting Program')