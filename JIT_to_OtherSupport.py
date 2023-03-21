import pandas as pd
import docx
from docx.shared import Pt
from docx.shared import Inches

# Read the CSV file into a pandas dataframe
df = pd.read_csv('JIT.csv', skiprows=6, header=6)


df = pd.DataFrame(df)
df = df[~df['PI (last, first)'].isin(['Active', 'Pending', 'Awarded'])]
df = df.reset_index(drop=True)
for i in range(len(df)):
    if df.loc[i, 'Is Mayo Secondary?'] == 'N':
        df.loc[i, 'Is Mayo Secondary?'] = 'Mayo Clinic, Rochester, MN'
    else:
        df.loc[i, 'Is Mayo Secondary?'] = '(sub award of ____)'

df['date_column'] = pd.to_datetime(df['Project Period End Date'], format='%m/%d/%Y')


def create_word_document():
    # Create a new Word document
    doc = docx.Document()

    # Set font and size
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)
    doc.styles['Normal'].paragraph_format.line_spacing = 1.0
    doc.styles['Normal'].paragraph_format.space_after = Pt(0)

    # Loop over grant_ID values
    for i in range(df.shape[0]):

        doc.add_paragraph("*Title: " + df["Grant Title"].loc[i])

        # Insert the rest of the text
        doc.add_paragraph("Major Goals: NEED ")
        doc.add_paragraph("*Status of Support: "+ df["Status Category"].loc[i])
        doc.add_paragraph("Project Number: " + df["External Grant ID"].astype(str).loc[i])
        doc.add_paragraph("Name of PD/PI: " + df["PI (last, first)"].loc[i])
        doc.add_paragraph("*Source of Support: "+ df["Funding Agency"].loc[i])
        doc.add_paragraph("*Primary Place of Performance: "+ df["Is Mayo Secondary?"].loc[i])
        doc.add_paragraph("Project/Proposal Start and End Date: (MM/YYYY) (if available): "+ df["Project Period Start Date"].loc[i] + " - " +df["Project Period End Date"].loc[i])
        doc.add_paragraph("* Total Award Amount (including Indirect Costs): "+ df["Total Project Period Costs (Direct plus Indirect)"].loc[i])
        doc.add_paragraph("* Person Months (Calendar/Academic/Summer) per budget period.")
        doc.add_paragraph()

        # Add a table to the document
        rows = df['# of Budget Periods in Current Project Period'].astype(int).loc[i]
        current_year = df['Current Year of Current Project Period'].astype(int).loc[i]
        table = doc.add_table(rows=rows +1, cols=2)
        table.style = 'Table Grid'
        doc.add_paragraph()

        df['year'] = df['date_column'].dt.year

                # Add year values to the first column in reverse order
        df['year'] = df['date_column'].dt.year
        for row in range(rows+1):
            for col in range(2):
                if col == 0:
                    if row == 0:
                        table.cell(row, 0).text = 'Year (YYYY)'
                    else:
                        table.cell(row, 0).text = str(row)+'. '+str(int(df['year'][i] - rows + row))
                if col ==1:
                    if row == 0:
                        table.cell(row, 1).text = 'Person Months (##.##)'
                    else:
                        table.cell(row, 1).text = df['Period {} Effort (Calendar months)'.format(row)].astype(str).loc[i]
						
        # Set the cell height and width
        for row in table.rows:
            for cell in row.cells:
                cell.width = Inches(1.81)
                cell.height = Inches(0.18)

        # Drop rows with NaN values
        table_rows = table.rows
        for row in reversed(range(1, rows + 1)):
            if any(cell.text == 'nan' for cell in table_rows[row].cells):
                table._element.remove(table_rows[row]._element)

    # Save the document
    doc.save('output1.docx')

create_word_document()
