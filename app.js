const express = require('express');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const multer = require('multer');
const exceljs = require('exceljs');

const app = express();

console.log(`process.env.ENV = ${process.env.ENV}`);
  mongoose.connect('mongodb://localhost:27017/bookApp', { useNewUrlParser: true, useUnifiedTopology: true });


const port = process.env.PORT || 3000;
const Book = require('./models/bookModel');
const bookRouter = require('./routes/bookRouter')(Book);

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());


const storage = multer.memoryStorage();
const upload = multer({ storage: storage });


app.post('/api/upload', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded.' });
  }

  try {
    const workbook = exceljs.read(req.file.buffer);
    const worksheet = workbook.getWorksheet(1); 

    const books = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber !== 1) {
        const [title, author, genre, read, publishYear, pagesCount, price] = row.values;

        const book = new Book({
          title,
          author,
          genre,
          read: read === 'true', 
          publishYear,
          pagesCount,
          price,
        });

        books.push(book);
      }
    });

    await Book.insertMany(books);

    return res.status(200).json({ message: 'File uploaded successfully.' });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: 'Error processing the file.' });
  }
});

app.get('/api/download', async (req, res) => {
  try {
    const books = await Book.find({}, { _id: 0, __v: 0 }); // Exclude _id and __v from the response

    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('Books');

    worksheet.addRow(['Title', 'Author', 'Genre', 'Read', 'Publish Year', 'Pages Count', 'Price']);

    books.forEach((book) => {
      const { title, author, genre, read, publishYear, pagesCount, price } = book;
      worksheet.addRow([title, author, genre, read, publishYear, pagesCount, price]);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    res.attachment('books.xlsx');
    res.send(buffer);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: 'Error fetching data for download.' });
  }
});

app.use('/api', bookRouter);

app.get('/', (req, res) => {
  res.send('Welcome to my Nodemon API!!');
});

const server = app.listen(port, () => {
  console.log(`Running on port ${port}`);
});

module.exports = server;
