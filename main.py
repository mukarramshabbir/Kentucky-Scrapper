import multiprocessing
import subprocess

def run_script(script_file):
    subprocess.run(['python', script_file])

if __name__ == '__main__':
    shared_queue = multiprocessing.Queue()
    script1_process = multiprocessing.Process(target=run_script, args=('Masachuset_scraper.py',))
    script1_process.start()
    
    script2_process = multiprocessing.Process(target=run_script, args=('Kentucky_Scraper.py',))

    
    script2_process.start()

    script1_process.join() 
    script2_process.join()
