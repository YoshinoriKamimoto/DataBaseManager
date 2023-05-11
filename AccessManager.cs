using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace WindowsFormsApp1
{
    // Accessファイルを操作するクラス
    internal class AccessManager
    {
        private OleDbConnection connection;
        private OleDbTransaction transaction = null;
        private string connectionStr;


        /// <summary>
        /// コンストラクタ
        /// </summary>
        public AccessManager(string connectStr)
        {
            // 接続文字列を代入
            this.connectionStr = connectStr;
        }


        /// <sumamry>
        /// Access接続メソッド
        /// </summary>
        public void OpenAccess()
        {
            try
            {
                // 接続用インスタンスを生成
                this.connection = new OleDbConnection(this.connectionStr);

                // 接続
                this.connection.Open();
            }
            catch
            {
                throw;
            }
        }


        /// <sumamry>
        /// Access切断メソッド
        /// </summary>
        public void CloseAccess()
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
                if (this.transaction != null)
                {
                    this.transaction.Commit();
                    this.transaction.Dispose();
                    this.transaction = null;
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
                if (this.transaction != null)
                {
                    this.transaction.Rollback();
                    this.transaction.Dispose();
                    this.transaction = null;
                }
            }
            catch
            {
                throw;
            }
        }


        /// <summary>
        /// SQL実行(戻り値：レコード数のみ)
        /// </summary>
        public int ExecuteNonQuery(string sqlStr)
        {
            try
            {
                
                OleDbCommand command;
                // トランザクション開始あり
                if (this.transaction != null)
                {
                    // SQL実行用インスタンスを生成
                    command = new OleDbCommand(sqlStr, this.connection, this.transaction);
                }
                else // トランザクション開始なし
                {
                    // SQL実行用インスタンスを生成
                    command = new OleDbCommand(sqlStr, this.connection);
                }

                // SQL実行
                int cnt = command.ExecuteNonQuery();
                return cnt;
            }
            catch
            {
                throw;
            }

        }


        /// <summary>
        /// SQL実行(戻り値:DataTable)
        /// </summary>
        public DataTable ExecuteQuery(string sqlStr)
        {
            try
            {
                // SQL実行用インスタンスを生成
                OleDbDataAdapter adapter = new OleDbDataAdapter(sqlStr, this.connection);

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
