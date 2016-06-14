class DownloadsController < ApplicationController
  def show
    send_file Xlsx.file(id, "#{id}.xlsx")
  end
end
