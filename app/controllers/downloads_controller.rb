class DownloadsController < ApplicationController
  def show
    id = params[:id]
    send_file Xlsx.file(id, "#{id}.xlsx")
  end
end
