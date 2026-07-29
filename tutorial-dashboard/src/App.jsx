import { Routes, Route } from 'react-router-dom';
import Home from './pages/Home';
import ToolTutorial from './pages/ToolTutorial';
import Layout from './components/Layout';

export default function App() {
  return (
    <Layout>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/tool/:toolId" element={<ToolTutorial />} />
      </Routes>
    </Layout>
  );
}
