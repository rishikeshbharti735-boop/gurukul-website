const express = require('express');
const mongoose = require('mongoose');
const path = require('path');
const Razorpay = require('razorpay');
const multer = require('multer');

const app = express();
const port = process.env.PORT || 3000; 

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Sahi rasta: static files 'school.html' folder ke andar hain
app.use(express.static(path.join(__dirname, 'school.html')));
app.use('/image', express.static(path.join(__dirname, 'image'))); 

// --- MONGODB ATLAS CONNECTION ---
const dbURI = 'mongodb+srv://rishikeshbharti735_db_user:0anl3dBSeFP4SmgC@cluster0.lohjt7p.mongodb.net/gurukulAashram?retryWrites=true&w=majority';

mongoose.connect(dbURI)
    .then(() => console.log('✅ MongoDB Atlas Connected!'))
    .catch(err => console.log('❌ MongoDB Atlas Error:', err));

// --- MULTER SETUP ---
const storage = multer.diskStorage({
    destination: './image/',
    filename: (req, file, cb) => {
        cb(null, Date.now() + '-' + file.originalname);
    }
});
const upload = multer({ storage: storage });

// --- RAZORPAY SETUP ---
const razorpay = new Razorpay({
    key_id: 'YOUR_KEY_ID', 
    key_secret: 'YOUR_KEY_SECRET'
});

// --- MODELS ---
const Contact = mongoose.model('Contact', new mongoose.Schema({ name: String, email: String, phone: String, subject: String, message: String }));
const Admission = mongoose.model('Admission', new mongoose.Schema({ student_name: String, dob: String, class_apply: String, parent_name: String, phone: String, address: String }));
const Notice = mongoose.model('Notice', new mongoose.Schema({ content: String, date: { type: Date, default: Date.now } }));
const Payment = mongoose.model('Payment', new mongoose.Schema({ student_name: String, amount: Number, payment_id: String, date: { type: Date, default: Date.now } }));
const Material = mongoose.model('Material', new mongoose.Schema({ className: String, link: String }));
const Gallery = mongoose.model('Gallery', new mongoose.Schema({ imagePath: String, caption: String, date: { type: Date, default: Date.now } }));
const Faculty = mongoose.model('Faculty', new mongoose.Schema({ name: String, designation: String, qualification: String, photo: String }));

// --- ROUTES ---

// 1. Home Page (school.html folder ke andar wala index.html)
app.get('/', (req, res) => { 
    res.sendFile(path.join(__dirname, 'school.html', 'index.html')); 
});

// Admin, Gallery, Faculty, aur Deletion routes (as specified in previous logic)
app.post('/admin-login', (req, res) => {
    const { username, password } = req.body;
    if (username === 'admin' && password === 'gurukul123') {
        res.json({ success: true });
    } else {
        res.json({ success: false, message: 'Galat Username ya Password!' });
    }
});

// Razorpay Order Creation
app.post('/create-order', async (req, res) => {
    try {
        const options = { amount: req.body.amount * 100, currency: "INR", receipt: "rcpt_" + Date.now() };
        const order = await razorpay.orders.create(options);
        res.json(order);
    } catch (error) { res.status(500).json({ error: "Order failed" }); }
});

app.post('/save-payment', async (req, res) => {
    try { await new Payment(req.body).save(); res.json({ success: true }); } 
    catch (e) { res.status(500).json({ success: false }); }
});

app.post('/update-material', async (req, res) => {
    try {
        await Material.findOneAndUpdate({ className: req.body.className }, { link: req.body.link }, { upsert: true });
        res.json({ success: true });
    } catch (e) { res.status(500).json({ success: false }); }
});

app.get('/get-materials', async (req, res) => {
    const materials = await Material.find({});
    res.json(materials);
});

app.post('/upload-photo', upload.single('schoolPhoto'), async (req, res) => {
    try {
        const newPhoto = new Gallery({ imagePath: '/image/' + req.file.filename, caption: req.body.caption });
        await newPhoto.save();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ success: false }); }
});

app.get('/get-gallery', async (req, res) => {
    const photos = await Gallery.find({}).sort({ date: -1 });
    res.json(photos);
});

app.post('/add-faculty', upload.single('teacherPhoto'), async (req, res) => {
    try {
        const newTeacher = new Faculty({
            name: req.body.name,
            designation: req.body.designation,
            qualification: req.body.qualification,
            photo: '/image/' + req.file.filename
        });
        await newTeacher.save();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ success: false }); }
});

app.get('/get-faculty', async (req, res) => {
    const faculty = await Faculty.find({});
    res.json(faculty);
});

app.post('/update-notice', async (req, res) => {
    try {
        await Notice.deleteMany({});
        await new Notice({ content: req.body.content }).save();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ success: false }); }
});

app.get('/get-notice', async (req, res) => {
    const notice = await Notice.findOne().sort({ date: -1 });
    res.json(notice || { content: "Welcome to Gurukul Aashram!" });
});

app.get('/admin-data', async (req, res) => {
    try {
        const admissions = await Admission.find({});
        const contacts = await Contact.find({});
        const payments = await Payment.find({}).sort({ date: -1 });
        const faculty = await Faculty.find({});
        res.json({ admissions, contacts, payments, faculty });
    } catch (e) { res.status(500).json({ error: "Error" }); }
});

app.post('/submit-admission', async (req, res) => {
    try { await new Admission(req.body).save(); res.send('<h2>Submitted! <a href="/">Go Back</a></h2>'); } 
    catch (e) { res.status(500).send('Error'); }
});

app.post('/submit-contact', async (req, res) => {
    try { await new Contact(req.body).save(); res.send('<h2>Sent! <a href="/">Go Back</a></h2>'); } 
    catch (e) { res.status(500).send('Error'); }
});

app.delete('/delete-admission/:id', async (req, res) => { await Admission.findByIdAndDelete(req.params.id); res.json({ success: true }); });
app.delete('/delete-contact/:id', async (req, res) => { await Contact.findByIdAndDelete(req.params.id); res.json({ success: true }); });
app.delete('/delete-faculty/:id', async (req, res) => { await Faculty.findByIdAndDelete(req.params.id); res.json({ success: true }); });

app.listen(port, () => { console.log(`🚀 Server: http://localhost:${port}`); });