// ============================================
// Firebase Configuration
// ============================================

const firebaseConfig = {
  apiKey: "AIzaSyCJG1AXPZ3aO_y9H7LVAL0mkxyv36O9Mx0",
  authDomain: "eiken-score-management.firebaseapp.com",
  projectId: "eiken-score-management",
  storageBucket: "eiken-score-management.firebasestorage.app",
  messagingSenderId: "391190530705",
  appId: "1:391190530705:web:af6963596243a4db404d30",
  measurementId: "G-T4KXSZRTK2"
};

// アクセスを許可するメールアドレスのリスト
// ここに登録されたGoogleアカウントのみデータにアクセスできます
const ALLOWED_EMAILS = [
  "kikuchi.mkt@gmail.com",
  "aizumiecc@gmail.com",
  "eccaizumi@gmail.com",
];

// セキュリティを高めるには、Firestoreルールにも同じメールアドレスを設定してください
