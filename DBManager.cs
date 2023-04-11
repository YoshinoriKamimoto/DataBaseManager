using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    // DBを操作するクラス(PostgreSQl,Npgsql使用)
    internal class DBManager
    {
        private NpgsqlConnection connection;
        private NpgsqlTransaction transaction;
        private string connectionStr;   


        /// <summary>
        /// コンストラクタ(外販DB用)
        /// </summary>
        public DBManager(string connectStr)
        {
            // 接続文字列を定義
            this.connectionStr = connectStr;
        }




        /// <summary>
        /// DB接続メソッド
        /// </summary>
        public void OpenDB()
        {
            try
            {
                // DB接続用インスタンスを生成
                this.connection = new NpgsqlConnection(this.connectionStr);

                // DB接続
                this.connection.Open();
            }
            catch
            {
                throw;
            }
            
        }


        /// <summary>
        /// DB切断メソッド
        /// </summary>
        public void CloseDB()
        {
            this.connection.Close();
            this.connection.Dispose();
        }



        /// <summary>
        /// トランザクション　開始メソッド
        /// </summary>
        public void BeginTran()
        {
            try
            {
                this.transaction = this.connection.BeginTransaction();
            }
            catch
            {
                throw;
            }
        }


        /// <summary>
        /// トランザクション　コミットメソッド
        /// </summary>
        public void CommitTran()
        {
            try
            {
                if (this.transaction.Connection != null)
                {
                    this.transaction.Commit();
                    this.transaction.Dispose();
                }
            }
            catch
            {
                throw;
            }
        }


        /// <summary>
        /// トランザクション　ロールバックメソッド
        /// </summary>
        public void RollBackTran()
        {
            try
            {
                if (this.transaction.Connection != null)
                {
                    this.transaction.Rollback();
                    this.transaction.Dispose();
                }
            }
            catch
            {
                throw;
            }
        }


        /// <summary>
        /// SQL実行(結果取得なし)
        /// </summary>
        public void ExecuteNonQuery(string sqlStr)
        {
            try
            {
                // SQL実行用インスタンスを生成
                NpgsqlCommand command = new NpgsqlCommand(sqlStr, this.connection);

                // SQL実行
                command.ExecuteNonQuery();
            }
            catch
            {
                throw;
            }
            
        }


        /// <summary>
        /// SQL実行(結果取得あり:DataTable)
        /// </summary>
        public DataTable ExecuteQuery(string sqlStr)
        {
            try
            {
                // SQL実行用インスタンスを生成
                NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sqlStr, this.connection);

                // SQL実行
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                // 返却
                return dt;
            }
            catch
            {
                throw;
            }
        }
    }
}
