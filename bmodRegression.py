#!/nfs/iil/proj/cto/dav/py3.6.3/bin/python3.6
import subprocess
import sys
import pandas as pd
import os
import argparse
from datetime import datetime

opt = argparse.ArgumentParser()
opt.add_argument('-x', '--xls', dest='xls', required=True,
                 help='Path to excel files contains the columns Cell, View ans CsvFile ', type=str,
                 default='~/bmodRegs/flc1278/BMOD_regression.xlsx')
opt.add_argument('-m', '--mailto', dest='maillist', required=False, help='list of UNIX usernames to send the mail to.',
                 nargs='+', type=str, default=os.environ['USER'])
args = opt.parse_args()
if type(args.maillist) != list:
    args.maillist = [args.maillist]

mail_massage = """To: {to}
Subject: Bmod Regression Results
Content-Type: text/html

<h2><FONT COLOR="BLUE"><strong>Schematis Status For BMOD:</strong></FONT></h2>
<p>Results Path: {rundir}</p>
{html_content}

"""


def color_format(val):
    if (val == 'fail'):
        return 'background-color: yellow'
    elif (val == 'pass'):
        return 'background-color: green'
    else:
        return ''


def mail_style(styler):
    styler.applymap(color_format)
    styler.set_table_styles([
        {"selector": "tbody tr:nth-child(even)",
         "props": [("background-color", "lightgrey")]},
        {"selector": "thead",
         "props": [("background-color", "yellowgreen")]},
        {"selector": "th",
         "props": [("padding", "2px 15px")]},
        {"selector": "td",
         "props": [("padding", "0 20px"), ("border-bottom:", "1pt solid grey")]},
    ])
    return styler


def run_command(param1, param2, param3, param4, rundir_name):
    command = '/nfs/iil/disks/falcon_18a_tc_rtl/users/yehudabe/work/analog_sa/verif/scripts/bmod_verifier.py -cell {} -view {} -csv {} -netlist {}/{}_bmodrun -nr > {}/{}_bmodrun.log'.format(
        param1, param3, param4, rundir_name, param1, rundir_name, param1)
    return subprocess.run(command, stderr=subprocess.PIPE, shell=True)


if __name__ == '__main__':

    # Get the path to the Excel file from the command line
    excel_file_path = args.xls
    if not os.path.exists(excel_file_path):
        print('Error: File not found at specified path')
    has_permissions = os.access(excel_file_path, os.R_OK | os.W_OK)
    if has_permissions:
        data_frame = pd.read_excel(excel_file_path)


        def strip_spaces(x):
            if isinstance(x, str):
                return x.strip()
            else:
                return x


        data_frame = data_frame.applymap(strip_spaces)  ## Remove whitw space

        # create run dir
        now = datetime.now().strftime("%m_%d_%Y__%H_%M")
        rundir_name = f'{os.environ["WORKAREA"]}/bmod_regression__{now}'
        if not os.path.exists(rundir_name):
            os.mkdir(rundir_name)

        # Extract the relevant information from the DataFrame
        param1 = data_frame['Cell'].values
        # param2 = data_frame['Lib'].values
        param3 = data_frame['View'].values
        param4 = data_frame['CsvFile'].values

        failed_commands = []
        error_messages = []
        status_list = []

        # Run the command for each row in the DataFrame
        for i in range(data_frame.shape[0]):
            print(f"Running Cell: {param1[i]} ")
            result = run_command(param1[i], 0, param3[i], param4[i], rundir_name)
            if result.returncode != 0:
                failed_commands.append(result.args)
                error_messages.append(result.stderr.decode())
            try:
                with open(f'{rundir_name}/{param1[i]}_bmodrun/res.dict', 'r') as res:
                    d = eval(res.read().split('\n')[0])
                    d.update({'Cell': param1[i]})
            except:
                print("Error : No res.dict found - skipping.. ")
                d = {'Cell': param1[i]}
            status_list.append(d)

        res_df = pd.DataFrame(status_list)
        col = res_df.columns
        col = col.tolist()[-1:] + col.tolist()[:-1]  # reordering
        res_df = res_df[col]
        print(res_df.to_markdown())
        print(f'Sending summary mail to : {args.maillist}')
        mailto = ','.join([f'{u}@ecsmtp.iil.intel.com' for u in args.maillist])
        with open(f'{rundir_name}/regression.mail', 'w') as f:
            m = mail_massage.format(to=mailto, rundir=rundir_name,
                                    html_content=res_df.style.pipe(mail_style)._repr_html_())
            f.write(m)
        subprocess.run(f"cat {rundir_name}/regression.mail | sendmail -t", stderr=subprocess.PIPE, shell=True)
        # Print the summary of failed commands
        if (len(failed_commands) > 0):
            print("Failed commands:")
            for i, cmd in enumerate(failed_commands):
                print("- Cmd command which failed: {}\n Error message: {}\n".format(cmd, error_messages[i]))
    else:
        print("The program does not have read and write permissions for the file or directory.")