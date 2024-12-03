import os

import pandas as pd
import yaml


def main():
    dir = "C:\\Users\\p.tkachenko\\Desktop\\mimir\\Mimir-main\\rules\\alerts"
    # dir = "C:\\Users\\paul\\Desktop\\Mimir-main\\rules\\alerts"
    files = os.listdir(dir)

    keys = ['alert', 'labels', 'expr', 'for', 'annotations', 'comment']
    result = []
    for file in files:
        ydir = f"{dir}\\{file}"
        with open(ydir, 'r') as alert:
            data = yaml.safe_load(alert)
        if data == None:
            continue

        for group in data['groups']:
            for rule in group['rules']:
                item = [file, group['name']]
                for key in keys:
                    try:
                        if key == 'labels':
                            item.append(rule[key]['severity'])
                        elif key == 'annotations':
                            item.append(rule[key]['summary'])
                            item.append(rule[key]['description'])
                        else:
                            val = rule[key]
                            item.append(val)
                    except KeyError as e:
                        item.append('Нет значения')
                        continue
                result.append(item)
                item = []

    df = pd.DataFrame(result, columns=['file', 'group', 'alert', 'severity', 'condition', 'delay', 'summary', 'description', 'comment'])

    writer = pd.ExcelWriter('output.xlsx')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.close()

if __name__ == '__main__':
    main()