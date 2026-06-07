import app from "./api/index.js";
import dotenv from "dotenv";

dotenv.config();

const PORT = process.env.PORT || 3050;

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
