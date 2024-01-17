const mongoose = require('mongoose');

const { Schema } = mongoose;

const bookModel = new Schema({
  title: { type: String, required: true, minlength: 2, maxlength: 30 },
  author: { type: String, required: true, minlength: 2, maxlength: 30 },
  genre: { type: String },
  read: { type: Boolean, default: false },
  publishYear: { type: Number, min: 1900, max: 2024 },
  pagesCount: { type: Number, min: 3, max: 1300 },
  price: { type: Number, min: 0, max: 150000 },
});

module.exports = mongoose.model('Book', bookModel);
