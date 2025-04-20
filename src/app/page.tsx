import StatementForm from '@/components/StatementForm';

export default function Home() {
  return (
    <main className="flex min-h-screen flex-col items-center justify-start p-8 bg-gray-900 text-gray-100">
      <h1 className="text-3xl font-bold mb-8">Bank Statement Analyzer</h1>
      <StatementForm />
    </main>
  );
}
