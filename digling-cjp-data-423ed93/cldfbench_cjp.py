import re
import pathlib

from clldutils.text import strip_chars
from cldfbench import Dataset as BaseDataset
from cldfbench import CLDFSpec

from poepy.poepy import Poems
import xlrd


class Dataset(BaseDataset):
    dir = pathlib.Path(__file__).parent
    id = "cjp"

    def cldf_specs(self):  # A dataset must declare all CLDF sets it creates.
        return CLDFSpec(dir=self.cldf_dir, module='Generic', metadata_fname='cldf-metadata.json')

    def cmd_download(self, args):
        pass

    def cmd_makecldf(self, args):
        args.writer.cldf.add_component('LanguageTable')
        #args.writer.cldf.add_component(
        #    'ExampleTable',
        #    'Text_ID',
        #    {'name': 'Sentence_Number', 'datatype': 'integer'},
        #    {'name': 'Phrase_Number', 'datatype': 'integer'},
        #)
        #args.writer.cldf.add_table('texts.csv', 'ID', 'Title')
        #args.writer.cldf.add_foreign_key('ExampleTable', 'Text_ID', 'texts.csv', 'ID')

        args.writer.objects['LanguageTable'].append({'ID': 'OldChinese', 'Name':
            'Old Chinese', 'Glottocode': 'oldc1244'})

        # read and test excel
        xlfile = xlrd.open_workbook(str(self.raw_dir.joinpath('cjp-data.xlsx')))
        sheet = xlfile.sheet_by_index(0)
        data = [dict(zip([h.lower() for h in sheet.row_values(3)], line)) for line in [
                    sheet.row_values(i) for i in range(4, sheet.nrows)]]
        print(data)

        # convert to Wordlist
        D = {0: [
            'poem',
            'edition',
            'stanza',
            'line_in_source',
            'line',
            'line_order',
            'rhymeids',
            'alignment',
            'source',
            'notes']}
        idx = 1
        for row in data:
            if row['id'].strip():
                print(row['id'])
                D[idx] = [row[h] for h in D[0]]
                # check for bad numbers and mark them zero
                rhymeids = []
                for number in row['rhymeids'].split():
                    if number.isdigit():
                        rhymeids += [int(number)]
                    else:
                        rhymeids += [0]
                        row['notes'] += ' Problematic rhymeids auto-corrected.'
                D[idx][D[0].index('rhymeids')] = rhymeids
                
                if len(row['alignment'].split(' + ')) != len(rhymeids):
                    print('problem', row['id'])
                    input()
                
                print(idx, row['rhymeids'])
                if idx != int(row['id']):
                    print(idx, row['id'])
                print(idx, row['line'])
                idx += 1
        
        poe = Poems(D, ref='rhymeids', fuzzy=True)
        poe._meta['poems'] = {'CJP': {}}
        poe.text(self.dir.joinpath('plots',
            'poems.html').as_posix(), 'CJP')



