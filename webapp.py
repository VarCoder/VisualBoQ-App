from main import removeFiles,Sheet,Agent
from pathlib import Path
from openpyxl import load_workbook
from flask import Flask, render_template, request, send_file

# Flask constructor
app = Flask(__name__)

# Root endpoint
@app.get('/')
def upload():
	return render_template('upload-excel.html')


@app.post('/view')
def view():

	# Read the File using Flask request
	file = request.files['file']
	# save file in local directory
	# file.save(file.filename)            
	dir = "tmp"
	xlPath = str(file.filename) # findExcelFile()
	removeFiles(Path(xlPath).with_suffix(".docx"), dir)
	print(f"Excel File Found: {xlPath}")

	wb = load_workbook(xlPath,
					read_only=True, data_only=True)

	infinity = Sheet("Infinity", 0)
	infinity.addColInfo(info="Modules", colStart="K", colEnd="M")
	infinity.addColInfo(info="Product", colStart="U")
	infinity.addColInfo(info="Space", colStart="D")
	infinity.addColInfo(info="Colors", colStart="O", colEnd="P")

	designer = Sheet("Designer", 1)
	designer.addColInfo(info="Modules", colStart="I", colEnd="N")
	designer.addColInfo(info="Switch", colStart="G")
	designer.addColInfo(info="Product", colStart="Y")
	designer.addColInfo(info="Space", colStart="D")
	designer.addColInfo(info="Colors", colStart="Q", colEnd="T")

	agent = Agent(wb, dir=dir, sheets=[
		infinity,
		designer
	], url="https://test.buildtrack.in/buildtrack/buildtrack-smart-switch/branches/buildtrack-smart-switch/app-src/")
	agent.openToIndia()
	agent.getModules()
	agent.getColors()
	agent.clickModules()
	agent.close()

	agent.publish(Path(xlPath).stem)
	# with open(agent.docx, "rb") as docx:
	# 	result = convert_to_html(docx)
	# Return HTML snippet that will render the table
	return send_file(agent.docx,as_attachment=True)


# Main Driver Function
if __name__ == '__main__':
	# Run the application on the local development server
	# app.run(debug=True)
    app.run(host="0.0.0.0", port=80)

