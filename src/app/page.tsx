import StatementForm from '@/components/StatementForm';

export default function Home() {
  return (
    <main className="flex min-h-screen flex-col items-center justify-start p-8 bg-gray-100">
      <h1 className="text-2xl font-semibold mb-6">Bank Statement Analyzer</h1>
      <StatementForm />
    </main>
  );
}
