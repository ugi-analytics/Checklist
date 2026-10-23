// Параметри проєкту Firebase (вставляються з консолі Firebase → Project settings → Your apps → Web app).
// Ці значення не є секретом: доступ до даних захищають правила firestore.rules.
export const FIREBASE_CONFIG = {
  apiKey: "AIzaSyAZ20Q4kjqFUJ4tM2qQiBHFoA0WWQgWpq4",
  authDomain: "palyvo-mvo.firebaseapp.com",
  projectId: "palyvo-mvo",
  storageBucket: "palyvo-mvo.firebasestorage.app",
  messagingSenderId: "739754266927",
  appId: "1:739754266927:web:e0e154f22ff41c115391da",
};
// E-mail, якому дозволено створити першого адміністратора (те саме значення — у firestore.rules).
export const ADMIN_EMAIL = "sbavinov1993@gmail.com";
