# admission-data-exporter
This is a utility program developed with nodejs. Using this program, you can export the information of the students who have applied to university of mazandaran. The program then uses the students' data and outputs a csv file with certain structure which can be used to import the information into the Golestan Learning Management System of University of Mazandaran. 
To get started, do the following tasks step by step: 

1. Get the list of the urls of excel files of applications from the website: https://admission.umz.ac.ir.
2. Get to the directory of the project with your commandline and install the required libraries with the command `npm install`
3. Open `index.html` and paste the urls of the applications in the main field you see on the page. Then press on the download button. Before this however, you need to have in mind that you need to have logged into your administration account at the admission website. This way, the server will let you download the files. otherwise you will see an error saying tha you are unauthorized. Next, click on the download files button to download all of the urls you had earlier pasted their into the field. 
4. Get the files that were stored in the `Downloads` folder of your computer and copy them into the `data-import` directory
5. run the program with `npm start`
6. You will have your exported csv file in the `data-export` folder. 

If you have any questions, contact me at zachshirow@gmail.com 
